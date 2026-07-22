using ClosedXML.Excel;
using Reports.Infrastructure.DTOs;
using Reports.Infrastructure.Exceptions;
using Reports.Infrastructure.Logger;
using Reports.Infrastructure.Models;
using Reports.Infrastructure.Repositories;
using System;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Printing;
using System.Text;
using System.Threading.Tasks;
using static Reports.Infrastructure.Models.Enums;

namespace Reports.Infrastructure.ReportGenerator
{
    public class ExcelReportGenerator : IReportGenerator
    {
        public Enums.ReportType Type => ReportType.Excel;
        private readonly IReportRepositoryAdoNet repositoryAdoNet;
        private readonly IReportRepositoryDapper repositoryDapper;
        private readonly ILogger logger;
        private readonly string libreOffice;


        public ExcelReportGenerator(IReportRepositoryAdoNet repositoryAdoNet, IReportRepositoryDapper repositoryDapper, ILogger logger, string libreOffice)
        {
            this.repositoryAdoNet = repositoryAdoNet;
            this.repositoryDapper = repositoryDapper;
            this.logger = logger;
            this.libreOffice = libreOffice;
        }

        public async Task<byte[]> ExecuteAsync(ReportRequest request, ReportDtl reportDtl)
        {
            try
            {

                Manifest manifest = await repositoryDapper.GetManifestByManifestIDAsync(request.ManifestID);

                //For SP parameters
                request.Parameters["ManifestID"] = request.ManifestID;

                DataSet dataSet = repositoryAdoNet.GetData(request, reportDtl);

                if (dataSet != null && dataSet.Tables.Count > 0 && dataSet.Tables[0].Rows.Count > 0)
                {
                    byte[] excel = GenerateExcel(dataSet, request, reportDtl, manifest);

                    // Legacy printing behavior — unchanged. OutputFormat has no effect here.
                    if (request.IsPrint)
                    {
                        if (request.PrinterName == "PDF")
                        {
                            byte[] excelToPdf = ConvertExcelToPdf(excel);
                            return excelToPdf;
                        }
                        PrintExcelDocument(excel, request.PrinterName);
                        return excel;
                    }

                    // Not printing (download/email flow): honor an explicit OutputFormat if the
                    // caller set one. Left null/empty => identical to previous behavior (raw xlsx).
                    if (!string.IsNullOrWhiteSpace(request.OutputFormat) &&
                        request.OutputFormat.Equals("PDF", StringComparison.OrdinalIgnoreCase))
                    {
                        return ConvertExcelToPdf(excel);
                    }

                    return excel;
                }
                else
                {
                    logger.WriteLog($"No data was returned for the report.");
                    return null;
                }
            }
            catch (CustomException ex)
            {
                logger.WriteLog($"Error to Execute Async: {ex}");
                throw;
            }
            catch (Exception ex)
            {
                logger.WriteLog($"Error to Execute Async: {ex}");
                throw new CustomException((int)ErrorMessages.ErrorCodes.GlobalError, ex.Message);
            }
        }
        public byte[] ConvertExcelToPdf(byte[] excelBytes)
        {
            string inputFilePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".xlsx");
            string outputFilePath = Path.ChangeExtension(inputFilePath, ".pdf");

            try
            {
                File.WriteAllBytes(inputFilePath, excelBytes);

                var conversionProcessStartInfo = new ProcessStartInfo
                {
                    FileName = libreOffice,
                    Arguments = $"--headless --nologo --norestore --convert-to pdf --outdir \"{Path.GetDirectoryName(outputFilePath)}\" \"{inputFilePath}\"",
                    RedirectStandardError = true,
                    RedirectStandardOutput = true,
                    UseShellExecute = false,
                    CreateNoWindow = true
                };

                using (Process process = new Process { StartInfo = conversionProcessStartInfo })
                {
                    process.Start();

                    string errorOutput = process.StandardError.ReadToEnd();
                    string standardOutput = process.StandardOutput.ReadToEnd();

                    process.WaitForExit();

                    if (process.ExitCode != 0 || !File.Exists(outputFilePath))
                    {
                        logger.WriteLog($"Failed to convert Excel to PDF. ExitCode: {process.ExitCode}");
                        if (!string.IsNullOrEmpty(errorOutput))
                        {
                            logger.WriteLog($"Error details: {errorOutput}");
                        }

                        if (!string.IsNullOrEmpty(standardOutput))
                        {
                            logger.WriteLog($"Error details: {standardOutput}");
                        }

                        throw new CustomException((int)ErrorMessages.ErrorCodes.GlobalError, "Failed to convert Excel to PDF.");
                    }

                    logger.WriteLog($"Excel to PDF conversion completed successfully.");

                    return File.ReadAllBytes(outputFilePath);

                }
            }
            catch (CustomException ex)
            {
                logger.WriteLog($"Failed to convert Excel to PDF: {ex}");
                throw;
            }
            catch (Exception ex)
            {
                logger.WriteLog($"Failed to convert Excel to PDF: {ex}");
                throw new CustomException((int)ErrorMessages.ErrorCodes.GlobalError, ex.Message);
            }
            finally
            {
                if (File.Exists(inputFilePath))
                    File.Delete(inputFilePath);

                if (File.Exists(outputFilePath))
                    File.Delete(outputFilePath);

            }
        }

        private void PrintExcelDocument(byte[] excel, string printerName)
        {
            string tempFilePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".xlsx");

            try
            {
                // Saving the Excel array to a temporary file
                File.WriteAllBytes(tempFilePath, excel);

                // Check the printer status
                PrintQueueStatus status = GetPrinterStatus(printerName);

                if (status.HasFlag(PrintQueueStatus.PaperProblem) ||
                status.HasFlag(PrintQueueStatus.Offline) ||
                status.HasFlag(PrintQueueStatus.Error))
                {
                    throw new CustomException((int)ErrorMessages.ErrorCodes.FailedToPrint, ErrorMessages.Messages[(int)ErrorMessages.ErrorCodes.FailedToPrint]);
                }

                // Preparing to run LibreOffice for printing
                var printProcess = new Process
                {
                    StartInfo = new ProcessStartInfo
                    {
                        FileName = libreOffice,
                        Arguments = $"--headless --nologo --norestore --pt \"{printerName}\" \"{tempFilePath}\"",
                        RedirectStandardError = true,
                        RedirectStandardOutput = true,
                        UseShellExecute = false,
                        CreateNoWindow = true
                    }
                };

                // Running the process (LibreOffice print)
                using (Process process = new Process { StartInfo = printProcess.StartInfo })
                {
                    process.Start();

                    string errorOutput = process.StandardError.ReadToEnd();
                    string standardOutput = process.StandardOutput.ReadToEnd();

                    process.WaitForExit();

                    if (process.ExitCode != 0)
                    {
                        logger.WriteLog($"Error to Print Excel Document. ExitCode: {process.ExitCode}");
                        if (!string.IsNullOrEmpty(errorOutput))
                        {
                            logger.WriteLog($"Error details: {errorOutput}");
                        }
                        throw new CustomException((int)ErrorMessages.ErrorCodes.FailedToPrint, ErrorMessages.Messages[(int)ErrorMessages.ErrorCodes.FailedToPrint]);
                    }
                    else
                    {
                        logger.WriteLog($"Print Excel Document completed successfully.");
                    }
                }
            }
            catch (CustomException ex)
            {
                logger.WriteLog($"Error to Print Excel Document: {ex}");
                throw;
            }
            catch (Exception ex)
            {
                logger.WriteLog($"Error to Print Excel Document: {ex}");
                throw new CustomException((int)ErrorMessages.ErrorCodes.FailedToPrint, $"{ErrorMessages.Messages[(int)ErrorMessages.ErrorCodes.FailedToPrint]} : {ex.Message}");
            }
            finally
            {
                // Deleting the temporary file
                if (File.Exists(tempFilePath))
                {
                    File.Delete(tempFilePath);
                }
            }
        }

        private PrintQueueStatus GetPrinterStatus(string printerName)
        {
            using (LocalPrintServer printServer = new LocalPrintServer())
            {
                foreach (PrintQueue queue in printServer.GetPrintQueues())
                {
                    if (queue.Name.Equals(printerName, StringComparison.OrdinalIgnoreCase))
                    {
                        queue.Refresh();
                        return queue.QueueStatus;
                    }
                }
            }
            throw new CustomException((int)ErrorMessages.ErrorCodes.UnknownPrinter, ErrorMessages.Messages[(int)ErrorMessages.ErrorCodes.UnknownPrinter]);
        }


        private byte[] GenerateExcel(DataSet dataSet, ReportRequest request, ReportDtl reportDtl, Manifest manifest)
        {
            try
            {
                if (Enum.TryParse(reportDtl.Template, true, out GenerateExcel generateExcel))
                {
                    switch (generateExcel)
                    {
                        case Enums.GenerateExcel.GenerateSUMentries9Report:
                            return GenerateSUMentries9Report(dataSet, request, reportDtl, manifest);
                        case Enums.GenerateExcel.GenerateInvBckReport:
                            return GenerateInvBckReport(dataSet, request, reportDtl, manifest);
                        case Enums.GenerateExcel.GenerateSUMvalindex3Report:
                            return GenerateSUMvalindex3Report(dataSet, request, reportDtl, manifest);
                        case Enums.GenerateExcel.GenerateSUMdeliveryGush8Report:
                            return GenerateSUMdeliveryGush8Report(dataSet, request, reportDtl, manifest);
                        case Enums.GenerateExcel.GenerateSUMdeliveryLines8Report:
                            return GenerateSUMdeliveryLines8Report(dataSet, request, reportDtl, manifest);
                        case Enums.GenerateExcel.GenerateCarsInShowroomsReport:
                            return GenerateCarsInShowroomsReport(dataSet, request, reportDtl, manifest);
                        case Enums.GenerateExcel.GenerateSUMqntIndex1Report:
                            return GenerateSUMqntIndex1Report(dataSet, request, reportDtl, manifest);
                        case Enums.GenerateExcel.GenerateDTLentries9Report:
                            return GenerateDTLentries9Report(dataSet, request, reportDtl, manifest);
                        case Enums.GenerateExcel.GenerateZeroInventory11Report:
                            return GenerateZeroInventory11Report(dataSet, request, reportDtl, manifest);
                        case Enums.GenerateExcel.GenerateContDtlReport:
                            return GenerateContDtlReport(dataSet, request, reportDtl, manifest);
                        case Enums.GenerateExcel.GeneratedeliveryCancel25Report:
                            return GeneratedeliveryCancel25Report(dataSet, request, reportDtl, manifest);
                        case Enums.GenerateExcel.GenerateRelease33Report:
                            return GenerateRelease33Report(dataSet, request, reportDtl, manifest);
                        case Enums.GenerateExcel.GenerateCustomerInvCover28Report:
                            return GenerateCustomerInvCover28Report(dataSet, request, reportDtl, manifest);
                        case Enums.GenerateExcel.GenerateStorageCalcReport:
                            return GenerateStorageCalcReport(dataSet, request, reportDtl, manifest);
                        case Enums.GenerateExcel.GenerateExtAutorityInv18Report:
                            return GenerateExtAutorityInv18Report(dataSet, request, reportDtl, manifest);
                        case Enums.GenerateExcel.GenerateCustomerInvAvg4Report:
                            return GenerateCustomerInvAvg4Report(dataSet, request, reportDtl, manifest);
                        case Enums.GenerateExcel.GenerateAVGStorageDays35Report:
                            return GenerateAVGStorageDays35Report(dataSet, request, reportDtl, manifest);
                        case Enums.GenerateExcel.GenerateSerialsReport:
                            return GenerateSerialsReport(dataSet, request, reportDtl, manifest);
                        case Enums.GenerateExcel.GenerateDelivery26Report:
                            return GenerateDelivery26Report(dataSet, request, reportDtl, manifest);
                        case Enums.GenerateExcel.GenerateReleaseGoods10Report:
                            return GenerateReleaseGoods10Report(dataSet, request, reportDtl, manifest);

                        default:
                            break;
                    }
                }

                return null;
            }
            catch (CustomException ex)
            {
                logger.WriteLog($"Error to Generate Excel: {ex}");
                throw;
            }
            catch (Exception ex)
            {
                logger.WriteLog($"Error to Generate Excel: {ex}");
                throw new CustomException((int)ErrorMessages.ErrorCodes.GlobalError, ex.Message);
            }
        }

        //Bold headings and unbold data
        private void ApplyTableStyleBoldHeadings(IXLTable tableRange)
        {
            tableRange.Theme = XLTableTheme.None;
            tableRange.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Right;
            tableRange.Style.Font.FontColor = XLColor.Black;
            tableRange.Style.Border.OutsideBorder = XLBorderStyleValues.None;

            var headerRow = tableRange.Range(1, 1, 1, tableRange.ColumnCount());
            headerRow.Style.Font.Bold = true;
            headerRow.Style.Border.BottomBorder = XLBorderStyleValues.Thin;

            foreach (var cell in headerRow.Cells())
            {
                cell.Style.Alignment.WrapText = true;
                cell.Value = cell.Value.ToString().Replace(" ", "\n");
            }

        }

        //Bold data and unbold headings
        private void ApplyTableStyleBoldData(IXLTable tableRange)
        {
            tableRange.Theme = XLTableTheme.None;
            tableRange.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Right;
            tableRange.Style.Font.FontColor = XLColor.Black;
            tableRange.Style.Border.OutsideBorder = XLBorderStyleValues.None;


            var headerRow = tableRange.Range(1, 1, 1, tableRange.ColumnCount());
            headerRow.Style.Font.Bold = false;
            headerRow.Style.Font.Underline = XLFontUnderlineValues.Single;


            var dataRange = tableRange.DataRange;
            dataRange.Style.Font.Bold = true;
        }

        private void AddHeader2(IXLWorksheet worksheet, ReportRequest request, ReportDtl reportDtl, Manifest manifest, int row, int numberOfColumns)
        {
            string manifestDetails = $"{manifest.GroupName}, בונדד {manifest.ManifestID} ({manifest.TermName})";


            string currentDateAndUser = $"{DateTime.Now.ToString("dd/MM/yyyy HH:mm")}, {request.User}";

            worksheet.Cell(row, 1).Value = manifestDetails;
            worksheet.Cell(row, 1).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Right;
            worksheet.Range(worksheet.Cell(row, 1), worksheet.Cell(row, 4)).Merge();


            worksheet.Cell(row, 5).Value = reportDtl.ReportName;
            worksheet.Cell(row, 5).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
            worksheet.Cell(row, 5).Style.Font.Bold = true;
            worksheet.Range(worksheet.Cell(row, 5), worksheet.Cell(row, numberOfColumns - 4)).Merge();


            worksheet.Cell(row, numberOfColumns - 3).Value = currentDateAndUser;
            worksheet.Cell(row, numberOfColumns - 3).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Left;
            worksheet.Range(worksheet.Cell(row, numberOfColumns - 3), worksheet.Cell(row, numberOfColumns)).Merge();

            row++;
            worksheet.Range(row, 1, row, numberOfColumns).Style.Border.TopBorder = XLBorderStyleValues.Thin;

            worksheet.Columns().AdjustToContents();
        }

        private void ApplyNumberFormatToSheet(IXLWorksheet worksheet)
        {
            foreach (var cell in worksheet.CellsUsed())
            {
                if (cell.DataType == XLDataType.Number || !string.IsNullOrEmpty(cell.FormulaA1))
                {
                    cell.Style.NumberFormat.Format = "#,##0";
                }
            }
        }

        private void PrintSettings(IXLWorksheet worksheet)
        {
            worksheet.PageSetup.PageOrientation = XLPageOrientation.Landscape;

            worksheet.RightToLeft = true;

            worksheet.Style.Font.FontSize = 10;

            // FitToPages(1, 0): "0" is an Excel-only convention meaning "unlimited pages tall".
            // LibreOffice (used by ConvertExcelToPdf and PrintExcelDocument for both the PDF
            // download/email path and physical printing) takes 0 literally and collapses the
            // whole report onto a single page instead. Use an explicit, generously large value
            // instead of 0 so width stays capped to 1 page while height is effectively unbounded
            // for any realistic report size.
            worksheet.PageSetup.FitToPages(1, 100);

            double marginSize = 0.5;
            worksheet.PageSetup.Margins.Left = marginSize;
            worksheet.PageSetup.Margins.Right = marginSize;
            worksheet.PageSetup.Margins.Top = marginSize;
            worksheet.PageSetup.Margins.Bottom = marginSize;

            worksheet.PageSetup.CenterHorizontally = true;
        }

        private void RemoveColumnsByName(DataTable table, List<string> columnsToRemove)
        {
            foreach (var columnName in columnsToRemove)
            {
                if (table.Columns.Contains(columnName))
                {
                    table.Columns.Remove(columnName);
                }
            }
        }

        private void ApplyColumnFormula(IXLWorksheet worksheet, DataTable dataTable, int startCalc, int endCalc, List<string> columnNames, string formulaType)
        {
            foreach (var columnName in columnNames)
            {
                if (!dataTable.Columns.Contains(columnName))
                    continue;

                int columnIndex = dataTable.Columns[columnName].Ordinal + 1;
                string columnLetter = XLHelper.GetColumnLetterFromNumber(columnIndex);
                string formula = $"{formulaType}({columnLetter}{startCalc}:{columnLetter}{endCalc - 1})";


                worksheet.Cell(endCalc, columnIndex).FormulaA1 = formula;
                worksheet.Cell(endCalc, columnIndex).Style.Font.Bold = true;
            }
        }

        private void InsertTableWithSubtotals(IXLWorksheet worksheet, DataTable table, string groupByColumn, List<string> columnsToSum, List<string> columnsToCount, int currentRow)
        {

            DataTable tableWithSubtotals = table.Clone();
            tableWithSubtotals.Columns.Add("IsSummaryRow", typeof(bool));

            var grouped = table.AsEnumerable().GroupBy(r => r[groupByColumn]);

            foreach (var group in grouped)
            {
                foreach (var row in group)
                {
                    var newRow = tableWithSubtotals.NewRow();
                    newRow.ItemArray = row.ItemArray;
                    newRow["IsSummaryRow"] = false;
                    tableWithSubtotals.Rows.Add(newRow);
                }

                // summaryRow for group
                var summaryRow = tableWithSubtotals.NewRow();

                foreach (var colName in columnsToSum)
                {
                    if (!tableWithSubtotals.Columns.Contains(colName))
                        continue;
                    var sum = group.Sum(r => decimal.TryParse(r[colName]?.ToString(), out var x) ? x : 0);
                    summaryRow[colName] = sum;
                }

                foreach (var colName in columnsToCount)
                {
                    if (!tableWithSubtotals.Columns.Contains(colName))
                        continue;
                    var count = group.Count();
                    summaryRow[colName] = count;
                }

                summaryRow["IsSummaryRow"] = true;
                tableWithSubtotals.Rows.Add(summaryRow);
            }

            // totalSummaryRow for table
            var totalRow = tableWithSubtotals.NewRow();

            foreach (var colName in columnsToSum)
            {
                if (!tableWithSubtotals.Columns.Contains(colName))
                    continue;
                var sum = table.AsEnumerable().Sum(r => decimal.TryParse(r[colName]?.ToString(), out var x) ? x : 0);
                totalRow[colName] = sum;
            }

            foreach (var colName in columnsToCount)
            {
                if (!tableWithSubtotals.Columns.Contains(colName))
                    continue;
                var count = table.Rows.Count;
                totalRow[colName] = count;
            }

            totalRow["IsSummaryRow"] = true;
            tableWithSubtotals.Rows.Add(totalRow);

            // insert table to sheet
            var tableRange = worksheet.Cell(currentRow, 1).InsertTable(tableWithSubtotals);
            ApplyTableStyleBoldHeadings(tableRange);

            int summaryColIndex = tableWithSubtotals.Columns["IsSummaryRow"].Ordinal + 1;

            foreach (var row in tableRange.Rows())
            {
                var cellValue = row.Cell(summaryColIndex).GetValue<string>();
                bool isSummary = false;
                bool.TryParse(cellValue, out isSummary);

                if (isSummary)
                {
                    row.Style.Font.Bold = true;
                    row.Style.Border.BottomBorder = XLBorderStyleValues.Thin;
                }
            }

            worksheet.Column(summaryColIndex).Delete();

        }

        private byte[] GenerateSUMentries9Report(DataSet dataSet, ReportRequest request, ReportDtl reportDtl, Manifest manifest)
        {
            try
            {
                using (var workbook = new XLWorkbook())
                {
                    var worksheet = workbook.Worksheets.Add(reportDtl.ReportID);

                    PrintSettings(worksheet);

                    int currentRow = 1;
                    int numberOfColumns = dataSet.Tables[0].Columns.Count;

                    AddHeader2(worksheet, request, reportDtl, manifest, currentRow, numberOfColumns);

                    currentRow += 2;

                    StringBuilder filter = new StringBuilder();

                    if (request.Parameters.ContainsKey("FromDate") && request.Parameters["FromDate"] != null && request.Parameters.ContainsKey("ToDate") && request.Parameters["ToDate"] != null)
                    {
                        string fromDate = request.Parameters["FromDate"]?.ToString();
                        string toDate = request.Parameters["ToDate"]?.ToString();
                        if (!string.IsNullOrEmpty(fromDate) && !string.IsNullOrEmpty(toDate)
                            && DateTime.TryParseExact(fromDate, "yyyy-MM-dd", CultureInfo.InvariantCulture, DateTimeStyles.None, out var parsedFromDate)
                            && DateTime.TryParseExact(toDate, "yyyy-MM-dd", CultureInfo.InvariantCulture, DateTimeStyles.None, out var parsedToDate))
                        {
                            filter.Append($"לתקופה: {parsedToDate:dd/MM/yy} - {parsedFromDate:dd/MM/yy}");
                        }
                    }

                    if (request.Parameters.ContainsKey("BilledImporterID") && request.Parameters["BilledImporterID"] != null)
                    {
                        string billedImporterID = request.Parameters["BilledImporterID"]?.ToString();
                        if (!string.IsNullOrEmpty(billedImporterID))
                        {
                            string billedImporter = dataSet.Tables[2].Rows[0]["BilledImporter"].ToString();
                            if (filter.Length > 0) filter.Append(", ");
                            filter.Append($"ללקוח: {billedImporter} ({billedImporterID})");
                        }
                    }

                    if (request.Parameters.ContainsKey("CustomsAgentID") && request.Parameters["CustomsAgentID"] != null)
                    {
                        string customsAgentID = request.Parameters["CustomsAgentID"]?.ToString();
                        if (!string.IsNullOrEmpty(customsAgentID))
                        {
                            string customsAgent = dataSet.Tables[3].Rows[0]["CustomsAgent"].ToString();
                            if (filter.Length > 0) filter.Append(", ");
                            filter.Append($"סוכן: {customsAgent} ({customsAgentID})");
                        }
                    }

                    if (request.Parameters.ContainsKey("CustomsAgentFile") && request.Parameters["CustomsAgentFile"] != null)
                    {
                        string customsAgentFile = request.Parameters["CustomsAgentFile"]?.ToString();
                        if (!string.IsNullOrEmpty(customsAgentFile))
                        {
                            if (filter.Length > 0) filter.Append(", ");
                            filter.Append($"תיק סוכן: {customsAgentFile}");
                        }
                    }

                    if (request.Parameters.ContainsKey("ConsignmentID") && request.Parameters["ConsignmentID"] != null)
                    {
                        string consignmentID = request.Parameters["ConsignmentID"]?.ToString();
                        if (!string.IsNullOrEmpty(consignmentID))
                        {
                            if (filter.Length > 0) filter.Append(", ");
                            filter.Append($"למזהה מטען: {consignmentID}");
                        }
                    }

                    if (request.Parameters.ContainsKey("FromGush") && request.Parameters["FromGush"] != null && request.Parameters.ContainsKey("ToGush") && request.Parameters["ToGush"] != null)
                    {
                        string fromGush = request.Parameters["FromGush"]?.ToString();
                        string toGush = request.Parameters["ToGush"]?.ToString();
                        if (!string.IsNullOrEmpty(fromGush) && !string.IsNullOrEmpty(toGush))
                        {
                            string FormattedFromGush = fromGush.Length > 2 ? $"{fromGush.Substring(0, 2)}/{fromGush.Substring(2)}" : fromGush;
                            string FormattedToGush = toGush.Length > 2 ? $"{toGush.Substring(0, 2)}/{toGush.Substring(2)}" : toGush;

                            if (filter.Length > 0) filter.Append(", ");
                            filter.Append($"מגוש {FormattedFromGush} עד גוש {FormattedToGush}");
                        }
                    }

                    if (request.Parameters.ContainsKey("DeliverySiteNumber") && request.Parameters["DeliverySiteNumber"] != null)
                    {
                        string deliverySiteNumber = request.Parameters["DeliverySiteNumber"]?.ToString();
                        if (!string.IsNullOrEmpty(deliverySiteNumber))
                        {
                            string name = dataSet.Tables[4].Rows[0]["Name"].ToString();
                            if (filter.Length > 0) filter.Append(", ");
                            filter.Append($"אתר מקור: {name}");
                        }
                    }

                    worksheet.Cell(currentRow, 1).Value = filter.ToString();
                    worksheet.Range(worksheet.Cell(currentRow, 1), worksheet.Cell(currentRow, numberOfColumns)).Merge();
                    worksheet.Cell(currentRow, 1).Style.Font.Bold = true;
                    worksheet.Cell(currentRow, 1).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Right;

                    currentRow += 2;

                    var table1 = worksheet.Cell(currentRow, 1).InsertTable(dataSet.Tables[0]);
                    ApplyTableStyleBoldHeadings(table1);


                    currentRow += dataSet.Tables[0].Rows.Count + 1;
                    worksheet.Cell(currentRow, 1).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Right;
                    worksheet.Range(currentRow, 1, currentRow, numberOfColumns).Style.Border.TopBorder = XLBorderStyleValues.Thin;

                    int startCalc = 6;
                    var columnsToSum = new List<string> { "כמות", "ערך בש\"ח", "20”", "40”", "נפח" };
                    var columnsToCount = new List<string> { "גוש" };
                    ApplyColumnFormula(worksheet, dataSet.Tables[0], startCalc, currentRow, columnsToSum, "SUM");
                    ApplyColumnFormula(worksheet, dataSet.Tables[0], startCalc, currentRow, columnsToCount, "COUNTA");

                    currentRow += 3;

                    worksheet.Cell(currentRow, 9).Value = "ריכוז כניסות לפי מטבע";
                    worksheet.Cell(currentRow, 9).Style.Font.Underline = XLFontUnderlineValues.Single;
                    worksheet.Cell(currentRow, 9).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Right;
                    worksheet.Cell(currentRow, 9).Style.Font.Bold = true;
                    worksheet.Range(worksheet.Cell(currentRow, 9), worksheet.Cell(currentRow, 11)).Merge();

                    currentRow++;

                    var table2 = worksheet.Cell(currentRow, 9).InsertTable(dataSet.Tables[1]);
                    ApplyTableStyleBoldData(table2);

                    ApplyNumberFormatToSheet(worksheet);

                    worksheet.Columns().AdjustToContents();


                    using (var stream = new MemoryStream())
                    {
                        workbook.SaveAs(stream);
                        return stream.ToArray();
                    }
                }
            }
            catch (Exception ex)
            {
                logger.WriteLog($"Error to Generate SUMentries9 Report: {ex}");
                throw new CustomException((int)ErrorMessages.ErrorCodes.GlobalError, ex.Message);
            }
        }

        private byte[] GenerateInvBckReport(DataSet dataSet, ReportRequest request, ReportDtl reportDtl, Manifest manifest)
        {
            try
            {
                using (var workbook = new XLWorkbook())
                {
                    var worksheet = workbook.Worksheets.Add(reportDtl.ReportID);

                    PrintSettings(worksheet);

                    if (request.IsPrint)
                    {
                        var columnsToDelete = new List<string> { "תיק סוכן", "כמות מוצהרת", "הערה" };
                        RemoveColumnsByName(dataSet.Tables[0], columnsToDelete);
                    }

                    int currentRow = 1;
                    int numberOfColumns = dataSet.Tables[0].Columns.Count;

                    AddHeader2(worksheet, request, reportDtl, manifest, currentRow, numberOfColumns);

                    currentRow += 2;

                    StringBuilder filter = new StringBuilder();

                    if (request.Parameters.ContainsKey("ValidityDate") && request.Parameters["ValidityDate"] != null)
                    {
                        string validityDate = request.Parameters["ValidityDate"]?.ToString();
                        if (!string.IsNullOrEmpty(validityDate) && DateTime.TryParseExact(validityDate, "yyyy-MM-dd", CultureInfo.InvariantCulture, DateTimeStyles.None, out var parsedDate))
                        {
                            filter.Append($"נכון לתאריך: {parsedDate:dd/MM/yy}");
                        }
                    }

                    if (request.Parameters.ContainsKey("BilledImporterID") && request.Parameters["BilledImporterID"] != null)
                    {
                        string billedImporterID = request.Parameters["BilledImporterID"]?.ToString();
                        if (!string.IsNullOrEmpty(billedImporterID))
                        {
                            string billedImporter = dataSet.Tables[1].Rows[0]["BilledImporter"].ToString();
                            if (filter.Length > 0) filter.Append(", ");
                            filter.Append($"ללקוח: {billedImporter} ({billedImporterID})");
                        }
                    }

                    if (request.Parameters.ContainsKey("BilledCAID") && request.Parameters["BilledCAID"] != null)
                    {
                        string billedCAID = request.Parameters["BilledCAID"]?.ToString();
                        if (!string.IsNullOrEmpty(billedCAID))
                        {
                            string billedCA = dataSet.Tables[2].Rows[0]["BilledCA"].ToString();
                            if (filter.Length > 0) filter.Append(", ");
                            filter.Append($"לסוכן: {billedCA} ({billedCAID})");
                        }
                    }

                    if (request.Parameters.ContainsKey("CustomsAgentFile") && request.Parameters["CustomsAgentFile"] != null)
                    {
                        string customsAgentFile = request.Parameters["CustomsAgentFile"]?.ToString();
                        if (!string.IsNullOrEmpty(customsAgentFile))
                        {
                            if (filter.Length > 0) filter.Append(", ");
                            filter.Append($"תיק: {customsAgentFile}");
                        }
                    }

                    if (request.Parameters.ContainsKey("FromGush") && request.Parameters["FromGush"] != null && request.Parameters.ContainsKey("ToGush") && request.Parameters["ToGush"] != null)
                    {
                        string fromGush = request.Parameters["FromGush"]?.ToString();
                        string toGush = request.Parameters["ToGush"]?.ToString();
                        if (!string.IsNullOrEmpty(fromGush) && !string.IsNullOrEmpty(toGush))
                        {
                            string FormattedFromGush = fromGush.Length > 2 ? $"{fromGush.Substring(0, 2)}/{fromGush.Substring(2)}" : fromGush;
                            string FormattedToGush = toGush.Length > 2 ? $"{toGush.Substring(0, 2)}/{toGush.Substring(2)}" : toGush;

                            if (filter.Length > 0) filter.Append(", ");
                            filter.Append($"מגוש {FormattedFromGush} עד גוש {FormattedToGush}");
                        }
                    }

                    if (request.Parameters.ContainsKey("GushState") && request.Parameters["GushState"] != null)
                    {
                        if (int.TryParse(request.Parameters["GushState"].ToString(), out int gushState))
                        {
                            if (filter.Length > 0) filter.Append(", ");

                            if (gushState == 0)
                            {
                                filter.Append("גושים פתוחים בלבד");
                            }
                            else if (gushState == 1)
                            {
                                filter.Append("גושים סגורים בלבד");
                            }
                        }
                    }

                    if (request.Parameters.ContainsKey("Location") && request.Parameters["Location"] != null)
                    {
                        string location = request.Parameters["Location"]?.ToString();
                        if (!string.IsNullOrEmpty(location))
                        {
                            if (filter.Length > 0) filter.Append(", ");
                            filter.Append($"איתור: {location}");
                        }
                    }

                    if (request.Parameters.ContainsKey("FromInv") && request.Parameters["FromInv"] != null && request.Parameters.ContainsKey("ToInv") && request.Parameters["ToInv"] != null)
                    {
                        string fromInv = Convert.ToInt32(request.Parameters["FromInv"]).ToString("N0");
                        string toInv = Convert.ToInt32(request.Parameters["ToInv"]).ToString("N0");
                        if (!string.IsNullOrEmpty(fromInv) && !string.IsNullOrEmpty(toInv))
                        {
                            if (filter.Length > 0) filter.Append(", ");
                            filter.Append($"מיתרה {fromInv} עד יתרה {toInv}");
                        }
                    }

                    if (request.Parameters.ContainsKey("Pcustoms") && request.Parameters["Pcustoms"] != null)
                    {
                        string pcustoms = request.Parameters["Pcustoms"]?.ToString();

                        if (filter.Length > 0) filter.Append(", ");

                        if (pcustoms == "VN")
                        {
                            filter.Append("לרכבים בלבד");
                        }
                        else if (pcustoms == "PP")
                        {
                            filter.Append("ללא רכבים בלבד");
                        }

                    }

                    worksheet.Cell(currentRow, 1).Value = filter.ToString();
                    worksheet.Range(worksheet.Cell(currentRow, 1), worksheet.Cell(currentRow, numberOfColumns)).Merge();
                    worksheet.Cell(currentRow, 1).Style.Font.Bold = true;
                    worksheet.Cell(currentRow, 1).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Right;

                    currentRow += 2;

                    var columnsToSum = new List<string> { "כמות מוצהרת", "טרם התקבל", "יתרה", "כמות משוחררת", "כמות ברשות מוסמכת" };
                    var columnsToCount = new List<string> { "גוש" };

                    if (request.IsPrint)
                    {
                        int mainSort = request.Parameters.TryGetValue("OrderBy", out var val) && val != null &&
                                       int.TryParse(val.ToString().FirstOrDefault().ToString(), out var result) ? result : 1;

                        var sortColumnMap = new Dictionary<int, string>
                        {
                            { 1, "שם הלקוח" },
                            { 2, "גוש" },
                            { 3, "הצהרה" },
                        };

                        string groupColumnName = sortColumnMap[mainSort];

                        InsertTableWithSubtotals(worksheet, dataSet.Tables[0], groupColumnName, columnsToSum, columnsToCount, currentRow);
                    }

                    else
                    {
                        var table1 = worksheet.Cell(currentRow, 1).InsertTable(dataSet.Tables[0]);
                        ApplyTableStyleBoldHeadings(table1);

                        currentRow += dataSet.Tables[0].Rows.Count + 1;
                        worksheet.Cell(currentRow, 1).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Right;
                        worksheet.Range(currentRow, 1, currentRow, numberOfColumns).Style.Border.TopBorder = XLBorderStyleValues.Thin;

                        int startCalc = 6;
                        ApplyColumnFormula(worksheet, dataSet.Tables[0], startCalc, currentRow, columnsToSum, "SUM");
                        ApplyColumnFormula(worksheet, dataSet.Tables[0], startCalc, currentRow, columnsToCount, "COUNTA");
                    }

                    ApplyNumberFormatToSheet(worksheet);

                    worksheet.Columns().AdjustToContents();

                    using (var stream = new MemoryStream())
                    {
                        workbook.SaveAs(stream);
                        return stream.ToArray();
                    }
                }
            }
            catch (Exception ex)
            {
                logger.WriteLog($"Error to Generate InvBck Report: {ex}");
                throw new CustomException((int)ErrorMessages.ErrorCodes.GlobalError, ex.Message);
            }
        }

        private byte[] GenerateSUMvalindex3Report(DataSet dataSet, ReportRequest request, ReportDtl reportDtl, Manifest manifest)
        {
            try
            {
                using (var workbook = new XLWorkbook())
                {
                    var worksheet = workbook.Worksheets.Add(reportDtl.ReportID);

                    PrintSettings(worksheet);

                    if (request.IsPrint)
                    {
                        var columnsToDelete = new List<string> { "סוכן מכס", "סוג אריזה כללי" };
                        RemoveColumnsByName(dataSet.Tables[0], columnsToDelete);
                    }

                    int currentRow = 1;
                    int numberOfColumns = dataSet.Tables[0].Columns.Count;

                    AddHeader2(worksheet, request, reportDtl, manifest, currentRow, numberOfColumns);

                    currentRow += 2;

                    StringBuilder filter = new StringBuilder();

                    if (request.Parameters.ContainsKey("ValidityDate") && request.Parameters["ValidityDate"] != null)
                    {
                        string validityDate = request.Parameters["ValidityDate"]?.ToString();
                        if (!string.IsNullOrEmpty(validityDate) && DateTime.TryParseExact(validityDate, "yyyy-MM-dd", CultureInfo.InvariantCulture, DateTimeStyles.None, out var parsedDate))
                        {
                            filter.Append($"נכון לתאריך: {parsedDate:dd/MM/yy}");
                        }
                    }

                    if (request.Parameters.ContainsKey("BilledImporterID") && request.Parameters["BilledImporterID"] != null)
                    {
                        string billedImporterID = request.Parameters["BilledImporterID"]?.ToString();
                        if (!string.IsNullOrEmpty(billedImporterID))
                        {
                            string billedImporter = dataSet.Tables[1].Rows[0]["BilledImporter"].ToString();
                            if (filter.Length > 0) filter.Append(", ");
                            filter.Append($"ללקוח: {billedImporter} ({billedImporterID})");
                        }
                    }

                    if (request.Parameters.ContainsKey("ConsignmentID") && request.Parameters["ConsignmentID"] != null)
                    {
                        string consignmentID = request.Parameters["ConsignmentID"]?.ToString();
                        if (!string.IsNullOrEmpty(consignmentID))
                        {
                            if (filter.Length > 0) filter.Append(", ");
                            filter.Append($"הצהרת אחסנה: {consignmentID}");
                        }
                    }

                    if (request.Parameters.ContainsKey("FromGush") && request.Parameters["FromGush"] != null && request.Parameters.ContainsKey("ToGush") && request.Parameters["ToGush"] != null)
                    {
                        string fromGush = request.Parameters["FromGush"]?.ToString();
                        string toGush = request.Parameters["ToGush"]?.ToString();
                        if (!string.IsNullOrEmpty(fromGush) && !string.IsNullOrEmpty(toGush))
                        {
                            string FormattedFromGush = fromGush.Length > 2 ? $"{fromGush.Substring(0, 2)}/{fromGush.Substring(2)}" : fromGush;
                            string FormattedToGush = toGush.Length > 2 ? $"{toGush.Substring(0, 2)}/{toGush.Substring(2)}" : toGush;

                            if (filter.Length > 0) filter.Append(", ");
                            filter.Append($"מגוש {FormattedFromGush} עד גוש {FormattedToGush}");
                        }
                    }
                    if (request.Parameters.ContainsKey("GushState") && request.Parameters["GushState"] != null)
                    {
                        if (int.TryParse(request.Parameters["GushState"].ToString(), out int gushState))
                        {
                            if (filter.Length > 0) filter.Append(", ");

                            if (gushState == 0)
                            {
                                filter.Append("גושים פתוחים בלבד");
                            }
                            else if (gushState == 1)
                            {
                                filter.Append("גושים סגורים בלבד");
                            }
                        }
                    }
                    if (request.Parameters.ContainsKey("ZeroInventory") && request.Parameters["ZeroInventory"] != null)
                    {
                        string zeroInventory = request.Parameters["ZeroInventory"]?.ToString();
                        if (!string.IsNullOrEmpty(zeroInventory) && zeroInventory == "Y")
                        {
                            if (filter.Length > 0) filter.Append(", ");
                            filter.Append("כולל יתרות אפס");
                        }
                    }

                    worksheet.Cell(currentRow, 1).Value = filter.ToString();
                    worksheet.Range(worksheet.Cell(currentRow, 1), worksheet.Cell(currentRow, numberOfColumns)).Merge();
                    worksheet.Cell(currentRow, 1).Style.Font.Bold = true;
                    worksheet.Cell(currentRow, 1).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Right;

                    currentRow += 2;

                    var table1 = worksheet.Cell(currentRow, 1).InsertTable(dataSet.Tables[0]);
                    ApplyTableStyleBoldHeadings(table1);


                    currentRow += dataSet.Tables[0].Rows.Count + 1;
                    worksheet.Cell(currentRow, 1).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Right;
                    worksheet.Range(currentRow, 1, currentRow, numberOfColumns).Style.Border.TopBorder = XLBorderStyleValues.Thin;

                    int startCalc = 6;
                    var columnsToSum = new List<string> { "כמות בפתיחה", "ערך בפתיחה", "נפח בפתיחה", "שטח בפתיחה", "משקל בפתיחה", "יתרת כמות", "יתרת ערך", "יתרת נפח", "יתרת שטח", "יתרת משקל" };
                    var columnsToCount = new List<string> { "גוש" };
                    ApplyColumnFormula(worksheet, dataSet.Tables[0], startCalc, currentRow, columnsToSum, "SUM");
                    ApplyColumnFormula(worksheet, dataSet.Tables[0], startCalc, currentRow, columnsToCount, "COUNTA");

                    ApplyNumberFormatToSheet(worksheet);

                    worksheet.Columns().AdjustToContents();

                    worksheet.Column(5).Width = 15;
                    worksheet.Column(6).Width = 15;

                    foreach (var row in worksheet.Rows())
                    {

                        row.Cell(5).Style.Alignment.WrapText = true;
                        row.Cell(6).Style.Alignment.WrapText = true;
                    }


                    using (var stream = new MemoryStream())
                    {
                        workbook.SaveAs(stream);
                        return stream.ToArray();
                    }
                }
            }
            catch (Exception ex)
            {
                logger.WriteLog($"Error to Generate SUMvalindex3 Report: {ex}");
                throw new CustomException((int)ErrorMessages.ErrorCodes.GlobalError, ex.Message);
            }
        }

        private byte[] GenerateSUMdeliveryGush8Report(DataSet dataSet, ReportRequest request, ReportDtl reportDtl, Manifest manifest)
        {
            try
            {
                using (var workbook = new XLWorkbook())
                {
                    var worksheet = workbook.Worksheets.Add(reportDtl.ReportID);

                    PrintSettings(worksheet);

                    int currentRow = 1;
                    int numberOfColumns = dataSet.Tables[0].Columns.Count;

                    AddHeader2(worksheet, request, reportDtl, manifest, currentRow, numberOfColumns);

                    currentRow += 2;

                    StringBuilder filter = new StringBuilder();

                    if (request.Parameters.ContainsKey("FromDate") && request.Parameters["FromDate"] != null && request.Parameters.ContainsKey("ToDate") && request.Parameters["ToDate"] != null)
                    {
                        string fromDate = request.Parameters["FromDate"]?.ToString();
                        string toDate = request.Parameters["ToDate"]?.ToString();
                        if (!string.IsNullOrEmpty(fromDate) && !string.IsNullOrEmpty(toDate)
                            && DateTime.TryParseExact(fromDate, "yyyy-MM-dd", CultureInfo.InvariantCulture, DateTimeStyles.None, out var parsedFromDate)
                            && DateTime.TryParseExact(toDate, "yyyy-MM-dd", CultureInfo.InvariantCulture, DateTimeStyles.None, out var parsedToDate))
                        {
                            filter.Append($"לתקופה: {parsedToDate:dd/MM/yy} - {parsedFromDate:dd/MM/yy}");
                        }
                    }

                    if (request.Parameters.ContainsKey("BilledImporterID") && request.Parameters["BilledImporterID"] != null)
                    {
                        string billedImporterID = request.Parameters["BilledImporterID"]?.ToString();
                        if (!string.IsNullOrEmpty(billedImporterID))
                        {
                            string billedImporter = dataSet.Tables[1].Rows[0]["BilledImporter"].ToString();
                            if (filter.Length > 0) filter.Append(", ");
                            filter.Append($"ללקוח: {billedImporter} ({billedImporterID})");
                        }
                    }


                    worksheet.Cell(currentRow, 1).Value = filter.ToString();
                    worksheet.Range(worksheet.Cell(currentRow, 1), worksheet.Cell(currentRow, numberOfColumns)).Merge();
                    worksheet.Cell(currentRow, 1).Style.Font.Bold = true;
                    worksheet.Cell(currentRow, 1).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Right;

                    currentRow += 2;

                    var table1 = worksheet.Cell(currentRow, 1).InsertTable(dataSet.Tables[0]);
                    ApplyTableStyleBoldHeadings(table1);


                    currentRow += dataSet.Tables[0].Rows.Count + 1;
                    worksheet.Cell(currentRow, 1).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Right;
                    worksheet.Range(currentRow, 1, currentRow, numberOfColumns).Style.Border.TopBorder = XLBorderStyleValues.Thin;

                    int startCalc = 6;
                    var columnsToSum = new List<string> { "כמות" };
                    var columnsToCount = new List<string> { "מונה" };
                    ApplyColumnFormula(worksheet, dataSet.Tables[0], startCalc, currentRow, columnsToSum, "SUM");
                    ApplyColumnFormula(worksheet, dataSet.Tables[0], startCalc, currentRow, columnsToCount, "COUNTA");

                    ApplyNumberFormatToSheet(worksheet);

                    worksheet.Columns().AdjustToContents();

                    using (var stream = new MemoryStream())
                    {
                        workbook.SaveAs(stream);
                        return stream.ToArray();
                    }
                }
            }
            catch (Exception ex)
            {
                logger.WriteLog($"Error to Generate SUMdeliveryGush8 Report: {ex}");
                throw new CustomException((int)ErrorMessages.ErrorCodes.GlobalError, ex.Message);
            }
        }

        private byte[] GenerateSUMdeliveryLines8Report(DataSet dataSet, ReportRequest request, ReportDtl reportDtl, Manifest manifest)
        {
            try
            {
                using (var workbook = new XLWorkbook())
                {
                    var worksheet = workbook.Worksheets.Add(reportDtl.ReportID);

                    PrintSettings(worksheet);

                    int currentRow = 1;
                    int numberOfColumns = dataSet.Tables[0].Columns.Count;

                    AddHeader2(worksheet, request, reportDtl, manifest, currentRow, numberOfColumns);

                    currentRow += 2;

                    StringBuilder filter = new StringBuilder();

                    if (request.Parameters.ContainsKey("FromDate") && request.Parameters["FromDate"] != null && request.Parameters.ContainsKey("ToDate") && request.Parameters["ToDate"] != null)
                    {
                        string fromDate = request.Parameters["FromDate"]?.ToString();
                        string toDate = request.Parameters["ToDate"]?.ToString();
                        if (!string.IsNullOrEmpty(fromDate) && !string.IsNullOrEmpty(toDate)
                            && DateTime.TryParseExact(fromDate, "yyyy-MM-dd", CultureInfo.InvariantCulture, DateTimeStyles.None, out var parsedFromDate)
                            && DateTime.TryParseExact(toDate, "yyyy-MM-dd", CultureInfo.InvariantCulture, DateTimeStyles.None, out var parsedToDate))
                        {
                            filter.Append($"לתקופה: {parsedToDate:dd/MM/yy} - {parsedFromDate:dd/MM/yy}");
                        }
                    }

                    if (request.Parameters.ContainsKey("BilledImporterID") && request.Parameters["BilledImporterID"] != null)
                    {
                        string billedImporterID = request.Parameters["BilledImporterID"]?.ToString();
                        if (!string.IsNullOrEmpty(billedImporterID))
                        {
                            string billedImporter = dataSet.Tables[1].Rows[0]["BilledImporter"].ToString();
                            string code = dataSet.Tables[1].Rows[0]["Code"].ToString();
                            if (filter.Length > 0) filter.Append(", ");
                            filter.Append($"ללקוח: {billedImporter} ({code})");
                        }
                    }

                    if (request.Parameters.ContainsKey("FromGush") && request.Parameters["FromGush"] != null && request.Parameters.ContainsKey("ToGush") && request.Parameters["ToGush"] != null)
                    {
                        string fromGush = request.Parameters["FromGush"]?.ToString();
                        string toGush = request.Parameters["ToGush"]?.ToString();
                        if (!string.IsNullOrEmpty(fromGush) && !string.IsNullOrEmpty(toGush))
                        {
                            string FormattedFromGush = fromGush.Length > 2 ? $"{fromGush.Substring(0, 2)}/{fromGush.Substring(2)}" : fromGush;
                            string FormattedToGush = toGush.Length > 2 ? $"{toGush.Substring(0, 2)}/{toGush.Substring(2)}" : toGush;

                            if (filter.Length > 0) filter.Append(", ");
                            filter.Append($"מגוש {FormattedFromGush} עד גוש {FormattedToGush}");
                        }
                    }

                    if (request.Parameters.ContainsKey("LineID") && request.Parameters["LineID"] != null)
                    {
                        string lineID = request.Parameters["LineID"]?.ToString();
                        if (!string.IsNullOrEmpty(lineID))
                        {
                            if (filter.Length > 0) filter.Append(", ");
                            filter.Append($"למזהה: {lineID}");
                        }
                    }


                    worksheet.Cell(currentRow, 1).Value = filter.ToString();
                    worksheet.Range(worksheet.Cell(currentRow, 1), worksheet.Cell(currentRow, numberOfColumns)).Merge();
                    worksheet.Cell(currentRow, 1).Style.Font.Bold = true;
                    worksheet.Cell(currentRow, 1).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Right;

                    currentRow += 2;

                    var table1 = worksheet.Cell(currentRow, 1).InsertTable(dataSet.Tables[0]);
                    ApplyTableStyleBoldHeadings(table1);


                    currentRow += dataSet.Tables[0].Rows.Count + 1;
                    worksheet.Cell(currentRow, 1).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Right;
                    worksheet.Range(currentRow, 1, currentRow, numberOfColumns).Style.Border.TopBorder = XLBorderStyleValues.Thin;

                    int startCalc = 6;
                    var columnsToSum = new List<string> { "כמות שנמסרה מהמזהה" };
                    var columnsToCount = new List<string> { "תעודת מסירה" };
                    ApplyColumnFormula(worksheet, dataSet.Tables[0], startCalc, currentRow, columnsToSum, "SUM");
                    ApplyColumnFormula(worksheet, dataSet.Tables[0], startCalc, currentRow, columnsToCount, "COUNTA");

                    ApplyNumberFormatToSheet(worksheet);

                    worksheet.Columns().AdjustToContents();

                    using (var stream = new MemoryStream())
                    {
                        workbook.SaveAs(stream);
                        return stream.ToArray();
                    }
                }
            }
            catch (Exception ex)
            {
                logger.WriteLog($"Error to Generate SUMdeliveryLines8 Report: {ex}");
                throw new CustomException((int)ErrorMessages.ErrorCodes.GlobalError, ex.Message);
            }
        }

        private byte[] GenerateCarsInShowroomsReport(DataSet dataSet, ReportRequest request, ReportDtl reportDtl, Manifest manifest)
        {
            try
            {
                using (var workbook = new XLWorkbook())
                {
                    var worksheet = workbook.Worksheets.Add(reportDtl.ReportID);

                    worksheet.PageSetup.PageOrientation = XLPageOrientation.Landscape;

                    worksheet.RightToLeft = true;

                    int currentRow = 1;
                    int numberOfColumns = dataSet.Tables[0].Columns.Count;

                    AddHeader2(worksheet, request, reportDtl, manifest, currentRow, numberOfColumns + 4);

                    currentRow += 2;

                    StringBuilder filter = new StringBuilder();

                    if (request.Parameters.ContainsKey("ValidityDate") && request.Parameters["ValidityDate"] != null)
                    {
                        string validityDate = request.Parameters["ValidityDate"]?.ToString();
                        if (!string.IsNullOrEmpty(validityDate) && DateTime.TryParseExact(validityDate, "yyyy-MM-dd", CultureInfo.InvariantCulture, DateTimeStyles.None, out var parsedDate))
                        {
                            filter.Append($"נכון לתאריך: {parsedDate:dd/MM/yy}");
                        }
                    }

                    if (request.Parameters.ContainsKey("BilledImporterID") && request.Parameters["BilledImporterID"] != null)
                    {
                        string billedImporterID = request.Parameters["BilledImporterID"]?.ToString();
                        if (!string.IsNullOrEmpty(billedImporterID))
                        {
                            string billedImporter = dataSet.Tables[1].Rows[0]["BilledImporter"].ToString();
                            if (filter.Length > 0) filter.Append(", ");
                            filter.Append($"ללקוח: {billedImporter} ({billedImporterID})");
                        }
                    }


                    worksheet.Cell(currentRow, 1).Value = filter.ToString();
                    worksheet.Range(worksheet.Cell(currentRow, 1), worksheet.Cell(currentRow, numberOfColumns + 4)).Merge();
                    worksheet.Cell(currentRow, 1).Style.Font.Bold = true;
                    worksheet.Cell(currentRow, 1).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Right;

                    currentRow += 2;

                    var table1 = worksheet.Cell(currentRow, 1).InsertTable(dataSet.Tables[0]);
                    ApplyTableStyleBoldHeadings(table1);

                    ApplyNumberFormatToSheet(worksheet);

                    worksheet.Columns().AdjustToContents();
                    worksheet.Column(1).Width += 5;

                    using (var stream = new MemoryStream())
                    {
                        workbook.SaveAs(stream);
                        return stream.ToArray();
                    }
                }
            }
            catch (Exception ex)
            {
                logger.WriteLog($"Error to Generate CarsInShowrooms Report: {ex}");
                throw new CustomException((int)ErrorMessages.ErrorCodes.GlobalError, ex.Message);
            }
        }

        private byte[] GenerateSUMqntIndex1Report(DataSet dataSet, ReportRequest request, ReportDtl reportDtl, Manifest manifest)
        {
            try
            {
                using (var workbook = new XLWorkbook())
                {
                    var worksheet = workbook.Worksheets.Add(reportDtl.ReportID);

                    PrintSettings(worksheet);

                    int currentRow = 1;
                    int numberOfColumns = dataSet.Tables[0].Columns.Count;

                    AddHeader2(worksheet, request, reportDtl, manifest, currentRow, numberOfColumns);

                    currentRow += 2;

                    StringBuilder filter = new StringBuilder();

                    if (request.Parameters.ContainsKey("ValidityDate") && request.Parameters["ValidityDate"] != null)
                    {
                        string validityDate = request.Parameters["ValidityDate"]?.ToString();
                        if (!string.IsNullOrEmpty(validityDate) && DateTime.TryParseExact(validityDate, "yyyy-MM-dd", CultureInfo.InvariantCulture, DateTimeStyles.None, out var parsedDate))
                        {
                            filter.Append($"נכון לתאריך: {parsedDate:dd/MM/yy}");
                        }
                    }

                    if (request.Parameters.ContainsKey("BilledImporterID") && request.Parameters["BilledImporterID"] != null)
                    {
                        string billedImporterID = request.Parameters["BilledImporterID"]?.ToString();
                        if (!string.IsNullOrEmpty(billedImporterID))
                        {
                            string billedImporter = dataSet.Tables[1].Rows[0]["BilledImporter"].ToString();
                            if (filter.Length > 0) filter.Append(", ");
                            filter.Append($"ללקוח: {billedImporter} ({billedImporterID})");
                        }
                    }

                    if (request.Parameters.ContainsKey("FromGush") && request.Parameters["FromGush"] != null && request.Parameters.ContainsKey("ToGush") && request.Parameters["ToGush"] != null)
                    {
                        string fromGush = request.Parameters["FromGush"]?.ToString();
                        string toGush = request.Parameters["ToGush"]?.ToString();
                        if (!string.IsNullOrEmpty(fromGush) && !string.IsNullOrEmpty(toGush))
                        {
                            string FormattedFromGush = fromGush.Length > 2 ? $"{fromGush.Substring(0, 2)}/{fromGush.Substring(2)}" : fromGush;
                            string FormattedToGush = toGush.Length > 2 ? $"{toGush.Substring(0, 2)}/{toGush.Substring(2)}" : toGush;

                            if (filter.Length > 0) filter.Append(", ");
                            filter.Append($"מגוש {FormattedFromGush} עד גוש {FormattedToGush}");
                        }
                    }

                    if (request.Parameters.ContainsKey("GushState") && request.Parameters["GushState"] != null)
                    {
                        if (int.TryParse(request.Parameters["GushState"].ToString(), out int gushState))
                        {
                            if (filter.Length > 0) filter.Append(", ");

                            if (gushState == 0)
                            {
                                filter.Append("גושים פתוחים בלבד");
                            }
                            else if (gushState == 1)
                            {
                                filter.Append("גושים סגורים בלבד");
                            }
                        }
                    }

                    if (request.Parameters.ContainsKey("FromInv") && request.Parameters["FromInv"] != null && request.Parameters.ContainsKey("ToInv") && request.Parameters["ToInv"] != null)
                    {
                        string fromInv = Convert.ToInt32(request.Parameters["FromInv"]).ToString("N0");
                        string toInv = Convert.ToInt32(request.Parameters["ToInv"]).ToString("N0");
                        if (!string.IsNullOrEmpty(fromInv) && !string.IsNullOrEmpty(toInv))
                        {
                            if (filter.Length > 0) filter.Append(", ");
                            filter.Append($"מיתרה {fromInv} עד יתרה {toInv}");
                        }
                    }

                    worksheet.Cell(currentRow, 1).Value = filter.ToString();
                    worksheet.Range(worksheet.Cell(currentRow, 1), worksheet.Cell(currentRow, numberOfColumns)).Merge();
                    worksheet.Cell(currentRow, 1).Style.Font.Bold = true;
                    worksheet.Cell(currentRow, 1).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Right;

                    currentRow += 2;

                    var table1 = worksheet.Cell(currentRow, 1).InsertTable(dataSet.Tables[0]);
                    ApplyTableStyleBoldHeadings(table1);


                    currentRow += dataSet.Tables[0].Rows.Count + 1;
                    worksheet.Cell(currentRow, 1).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Right;
                    worksheet.Range(currentRow, 1, currentRow, numberOfColumns).Style.Border.TopBorder = XLBorderStyleValues.Thin;

                    int startCalc = 6;
                    var columnsToSum = new List<string> { "יתרת כמות" };
                    var columnsToCount = new List<string> { "גוש" };
                    ApplyColumnFormula(worksheet, dataSet.Tables[0], startCalc, currentRow, columnsToSum, "SUM");
                    ApplyColumnFormula(worksheet, dataSet.Tables[0], startCalc, currentRow, columnsToCount, "COUNTA");


                    ApplyNumberFormatToSheet(worksheet);

                    worksheet.Columns().AdjustToContents();

                    worksheet.Column(4).Width = 25;
                    worksheet.Column(5).Width = 25;
                    worksheet.Column(6).Width = 25;
                    worksheet.Column(13).Width = 25;

                    foreach (var row in worksheet.Rows())
                    {
                        row.Cell(4).Style.Alignment.WrapText = true;
                        row.Cell(5).Style.Alignment.WrapText = true;
                        row.Cell(6).Style.Alignment.WrapText = true;
                        row.Cell(13).Style.Alignment.WrapText = true;
                    }


                    using (var stream = new MemoryStream())
                    {
                        workbook.SaveAs(stream);
                        return stream.ToArray();
                    }
                }
            }
            catch (Exception ex)
            {
                logger.WriteLog($"Error to Generate SUMqntIndex1 Report: {ex}");
                throw new CustomException((int)ErrorMessages.ErrorCodes.GlobalError, ex.Message);
            }
        }

        private byte[] GenerateDTLentries9Report(DataSet dataSet, ReportRequest request, ReportDtl reportDtl, Manifest manifest)
        {
            try
            {
                using (var workbook = new XLWorkbook())
                {
                    var worksheet = workbook.Worksheets.Add(reportDtl.ReportID);

                    PrintSettings(worksheet);

                    int currentRow = 1;
                    int numberOfColumns = dataSet.Tables[0].Columns.Count;

                    AddHeader2(worksheet, request, reportDtl, manifest, currentRow, numberOfColumns);

                    currentRow += 2;

                    StringBuilder filter = new StringBuilder();

                    if (request.Parameters.ContainsKey("FromDate") && request.Parameters["FromDate"] != null && request.Parameters.ContainsKey("ToDate") && request.Parameters["ToDate"] != null)
                    {
                        string fromDate = request.Parameters["FromDate"]?.ToString();
                        string toDate = request.Parameters["ToDate"]?.ToString();
                        if (!string.IsNullOrEmpty(fromDate) && !string.IsNullOrEmpty(toDate)
                            && DateTime.TryParseExact(fromDate, "yyyy-MM-dd", CultureInfo.InvariantCulture, DateTimeStyles.None, out var parsedFromDate)
                            && DateTime.TryParseExact(toDate, "yyyy-MM-dd", CultureInfo.InvariantCulture, DateTimeStyles.None, out var parsedToDate))
                        {
                            filter.Append($"לתקופה: {parsedToDate:dd/MM/yy} - {parsedFromDate:dd/MM/yy}");
                        }
                    }

                    if (request.Parameters.ContainsKey("BilledImporterID") && request.Parameters["BilledImporterID"] != null)
                    {
                        string billedImporterID = request.Parameters["BilledImporterID"]?.ToString();
                        if (!string.IsNullOrEmpty(billedImporterID))
                        {
                            string billedImporter = dataSet.Tables[1].Rows[0]["BilledImporter"].ToString();
                            if (filter.Length > 0) filter.Append(", ");
                            filter.Append($"ללקוח: {billedImporter} ({billedImporterID})");
                        }
                    }

                    if (request.Parameters.ContainsKey("CustomsAgentID") && request.Parameters["CustomsAgentID"] != null)
                    {
                        string customsAgentID = request.Parameters["CustomsAgentID"]?.ToString();
                        if (!string.IsNullOrEmpty(customsAgentID))
                        {
                            string customsAgent = dataSet.Tables[2].Rows[0]["CustomsAgent"].ToString();
                            if (filter.Length > 0) filter.Append(", ");
                            filter.Append($"סוכן: {customsAgent} ({customsAgentID})");
                        }
                    }

                    if (request.Parameters.ContainsKey("CustomsAgentFile") && request.Parameters["CustomsAgentFile"] != null)
                    {
                        string customsAgentFile = request.Parameters["CustomsAgentFile"]?.ToString();
                        if (!string.IsNullOrEmpty(customsAgentFile))
                        {
                            if (filter.Length > 0) filter.Append(", ");
                            filter.Append($"תיק סוכן: {customsAgentFile}");
                        }
                    }

                    if (request.Parameters.ContainsKey("ConsignmentID") && request.Parameters["ConsignmentID"] != null)
                    {
                        string consignmentID = request.Parameters["ConsignmentID"]?.ToString();
                        if (!string.IsNullOrEmpty(consignmentID))
                        {
                            if (filter.Length > 0) filter.Append(", ");
                            filter.Append($"מטען: {consignmentID}");
                        }
                    }

                    if (request.Parameters.ContainsKey("FromGush") && request.Parameters["FromGush"] != null && request.Parameters.ContainsKey("ToGush") && request.Parameters["ToGush"] != null)
                    {
                        string fromGush = request.Parameters["FromGush"]?.ToString();
                        string toGush = request.Parameters["ToGush"]?.ToString();
                        if (!string.IsNullOrEmpty(fromGush) && !string.IsNullOrEmpty(toGush))
                        {
                            string FormattedFromGush = fromGush.Length > 2 ? $"{fromGush.Substring(0, 2)}/{fromGush.Substring(2)}" : fromGush;
                            string FormattedToGush = toGush.Length > 2 ? $"{toGush.Substring(0, 2)}/{toGush.Substring(2)}" : toGush;

                            if (filter.Length > 0) filter.Append(", ");
                            filter.Append($"מגוש {FormattedFromGush} עד גוש {FormattedToGush}");
                        }
                    }

                    if (request.Parameters.ContainsKey("LineID") && request.Parameters["LineID"] != null)
                    {
                        string lineID = request.Parameters["LineID"]?.ToString();
                        if (!string.IsNullOrEmpty(lineID))
                        {
                            if (filter.Length > 0) filter.Append(", ");
                            filter.Append($"מזהה: {lineID}");
                        }
                    }

                    if (request.Parameters.ContainsKey("DeliverySiteNumber") && request.Parameters["DeliverySiteNumber"] != null)
                    {
                        string deliverySiteNumber = request.Parameters["DeliverySiteNumber"]?.ToString();
                        if (!string.IsNullOrEmpty(deliverySiteNumber))
                        {
                            string name = dataSet.Tables[3].Rows[0]["Name"].ToString();
                            if (filter.Length > 0) filter.Append(", ");
                            filter.Append($"אתר מקור: {name}");
                        }
                    }

                    worksheet.Cell(currentRow, 1).Value = filter.ToString();
                    worksheet.Range(worksheet.Cell(currentRow, 1), worksheet.Cell(currentRow, numberOfColumns)).Merge();
                    worksheet.Cell(currentRow, 1).Style.Font.Bold = true;
                    worksheet.Cell(currentRow, 1).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Right;
                    worksheet.Cell(currentRow, 1).Style.Alignment.WrapText = true;
                    worksheet.Row(currentRow).Height = 25;

                    currentRow += 2;

                    var table1 = worksheet.Cell(currentRow, 1).InsertTable(dataSet.Tables[0]);
                    ApplyTableStyleBoldHeadings(table1);


                    currentRow += dataSet.Tables[0].Rows.Count + 1;
                    worksheet.Cell(currentRow, 1).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Right;
                    worksheet.Range(currentRow, 1, currentRow, numberOfColumns).Style.Border.TopBorder = XLBorderStyleValues.Thin;

                    int startCalc = 6;
                    var columnsToSum = new List<string> { "כמות" };
                    ApplyColumnFormula(worksheet, dataSet.Tables[0], startCalc, currentRow, columnsToSum, "SUM");

                    ApplyNumberFormatToSheet(worksheet);

                    worksheet.Columns().AdjustToContents();


                    using (var stream = new MemoryStream())
                    {
                        workbook.SaveAs(stream);
                        return stream.ToArray();
                    }
                }
            }
            catch (Exception ex)
            {
                logger.WriteLog($"Error to Generate DTLentries9 Report: {ex}");
                throw new CustomException((int)ErrorMessages.ErrorCodes.GlobalError, ex.Message);
            }
        }

        private byte[] GenerateZeroInventory11Report(DataSet dataSet, ReportRequest request, ReportDtl reportDtl, Manifest manifest)
        {
            try
            {
                using (var workbook = new XLWorkbook())
                {
                    var worksheet = workbook.Worksheets.Add(reportDtl.ReportID);

                    PrintSettings(worksheet);

                    int currentRow = 1;
                    int numberOfColumns = dataSet.Tables[0].Columns.Count;

                    AddHeader2(worksheet, request, reportDtl, manifest, currentRow, numberOfColumns + 3);

                    currentRow += 2;

                    StringBuilder filter = new StringBuilder();

                    if (request.Parameters.ContainsKey("ZeroInvDate") && request.Parameters["ZeroInvDate"] != null)
                    {
                        string zeroInvDate = request.Parameters["ZeroInvDate"]?.ToString();
                        if (!string.IsNullOrEmpty(zeroInvDate)
                            && DateTime.TryParseExact(zeroInvDate, "yyyy-MM-dd", CultureInfo.InvariantCulture, DateTimeStyles.None, out var parsedZeroInvDate))
                        {
                            filter.Append($"גושים שיתרתם התאפסה החל מ: {parsedZeroInvDate:dd/MM/yy}");
                        }
                    }

                    if (request.Parameters.ContainsKey("BilledImporterID") && request.Parameters["BilledImporterID"] != null)
                    {
                        string billedImporterID = request.Parameters["BilledImporterID"]?.ToString();
                        if (!string.IsNullOrEmpty(billedImporterID))
                        {
                            string billedImporter = dataSet.Tables[1].Rows[0]["BilledImporter"].ToString();
                            if (filter.Length > 0) filter.Append(", ");
                            filter.Append($"ללקוח: {billedImporter} ({billedImporterID})");
                        }
                    }

                    if (request.Parameters.ContainsKey("FromGush") && request.Parameters["FromGush"] != null && request.Parameters.ContainsKey("ToGush") && request.Parameters["ToGush"] != null)
                    {
                        string fromGush = request.Parameters["FromGush"]?.ToString();
                        string toGush = request.Parameters["ToGush"]?.ToString();
                        if (!string.IsNullOrEmpty(fromGush) && !string.IsNullOrEmpty(toGush))
                        {
                            string FormattedFromGush = fromGush.Length > 2 ? $"{fromGush.Substring(0, 2)}/{fromGush.Substring(2)}" : fromGush;
                            string FormattedToGush = toGush.Length > 2 ? $"{toGush.Substring(0, 2)}/{toGush.Substring(2)}" : toGush;

                            if (filter.Length > 0) filter.Append(", ");
                            filter.Append($"מגוש {FormattedFromGush} עד גוש {FormattedToGush}");
                        }
                    }

                    if (request.Parameters.ContainsKey("AlreadyClosed") && request.Parameters["AlreadyClosed"] != null)
                    {
                        string alreadyClosed = request.Parameters["AlreadyClosed"]?.ToString();
                        if (!string.IsNullOrEmpty(alreadyClosed) && alreadyClosed == "Y")
                        {
                            if (filter.Length > 0) filter.Append(", ");
                            filter.Append("כולל גושים סגורים");
                        }
                    }

                    worksheet.Cell(currentRow, 1).Value = filter.ToString();
                    worksheet.Range(worksheet.Cell(currentRow, 1), worksheet.Cell(currentRow, numberOfColumns + 3)).Merge();
                    worksheet.Cell(currentRow, 1).Style.Font.Bold = true;
                    worksheet.Cell(currentRow, 1).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Right;

                    currentRow += 2;

                    var table1 = worksheet.Cell(currentRow, 1).InsertTable(dataSet.Tables[0]);
                    ApplyTableStyleBoldHeadings(table1);


                    currentRow += dataSet.Tables[0].Rows.Count + 1;
                    worksheet.Cell(currentRow, 1).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Right;
                    worksheet.Range(currentRow, 1, currentRow, numberOfColumns).Style.Border.TopBorder = XLBorderStyleValues.Thin;


                    ApplyNumberFormatToSheet(worksheet);

                    worksheet.Columns().AdjustToContents();

                    worksheet.Column(5).Width = 10;
                    worksheet.Column(6).Width = 10;


                    using (var stream = new MemoryStream())
                    {
                        workbook.SaveAs(stream);
                        return stream.ToArray();
                    }
                }
            }
            catch (Exception ex)
            {
                logger.WriteLog($"Error to Generate ZeroInventory11 Report: {ex}");
                throw new CustomException((int)ErrorMessages.ErrorCodes.GlobalError, ex.Message);
            }
        }

        private byte[] GenerateContDtlReport(DataSet dataSet, ReportRequest request, ReportDtl reportDtl, Manifest manifest)
        {
            try
            {
                using (var workbook = new XLWorkbook())
                {
                    var worksheet = workbook.Worksheets.Add(reportDtl.ReportID);

                    PrintSettings(worksheet);

                    int currentRow = 1;
                    int numberOfColumns = dataSet.Tables[0].Columns.Count;

                    AddHeader2(worksheet, request, reportDtl, manifest, currentRow, numberOfColumns);

                    currentRow += 2;

                    StringBuilder filter = new StringBuilder();

                    if (request.Parameters.ContainsKey("BilledImporterID") && request.Parameters["BilledImporterID"] != null)
                    {
                        string billedImporterID = request.Parameters["BilledImporterID"]?.ToString();
                        if (!string.IsNullOrEmpty(billedImporterID))
                        {
                            string billedImporter = dataSet.Tables[1].Rows[0]["BilledImporter"].ToString();
                            filter.Append($"ללקוח: {billedImporter} ({billedImporterID})");
                        }
                    }

                    if (request.Parameters.ContainsKey("FromGush") && request.Parameters["FromGush"] != null && request.Parameters.ContainsKey("ToGush") && request.Parameters["ToGush"] != null)
                    {
                        string fromGush = request.Parameters["FromGush"]?.ToString();
                        string toGush = request.Parameters["ToGush"]?.ToString();
                        if (!string.IsNullOrEmpty(fromGush) && !string.IsNullOrEmpty(toGush))
                        {
                            string FormattedFromGush = fromGush.Length > 2 ? $"{fromGush.Substring(0, 2)}/{fromGush.Substring(2)}" : fromGush;
                            string FormattedToGush = toGush.Length > 2 ? $"{toGush.Substring(0, 2)}/{toGush.Substring(2)}" : toGush;

                            if (filter.Length > 0) filter.Append(", ");
                            filter.Append($"מגוש {FormattedFromGush} עד גוש {FormattedToGush}");
                        }
                    }

                    if (request.Parameters.ContainsKey("FromDateIn") && request.Parameters["FromDateIn"] != null && request.Parameters.ContainsKey("ToDateIn") && request.Parameters["ToDateIn"] != null)
                    {
                        string fromDateIn = request.Parameters["FromDateIn"]?.ToString();
                        string toDateIn = request.Parameters["ToDateIn"]?.ToString();
                        if (!string.IsNullOrEmpty(fromDateIn) && !string.IsNullOrEmpty(toDateIn)
                            && DateTime.TryParseExact(fromDateIn, "yyyy-MM-dd", CultureInfo.InvariantCulture, DateTimeStyles.None, out var parsedFromDateIn)
                            && DateTime.TryParseExact(toDateIn, "yyyy-MM-dd", CultureInfo.InvariantCulture, DateTimeStyles.None, out var parsedToDateIn))
                        {
                            if (filter.Length > 0) filter.Append(", ");
                            filter.Append($"תאריך כניסה {parsedFromDateIn:dd/MM/yy} עד {parsedToDateIn:dd/MM/yy}");
                        }
                    }

                    if (request.Parameters.ContainsKey("FromDateMT") && request.Parameters["FromDateMT"] != null && request.Parameters.ContainsKey("ToDateMT") && request.Parameters["ToDateMT"] != null)
                    {
                        string fromDateMT = request.Parameters["FromDateMT"]?.ToString();
                        string toDateMT = request.Parameters["ToDateMT"]?.ToString();
                        if (!string.IsNullOrEmpty(fromDateMT) && !string.IsNullOrEmpty(toDateMT)
                            && DateTime.TryParseExact(fromDateMT, "yyyy-MM-dd", CultureInfo.InvariantCulture, DateTimeStyles.None, out var parsedFromDateMT)
                            && DateTime.TryParseExact(toDateMT, "yyyy-MM-dd", CultureInfo.InvariantCulture, DateTimeStyles.None, out var parsedToDateMT))
                        {
                            if (filter.Length > 0) filter.Append(", ");
                            filter.Append($"מתאריך ריקון {parsedFromDateMT:dd/MM/yy} עד {parsedToDateMT:dd/MM/yy}");
                        }
                    }

                    if (request.Parameters.ContainsKey("FromDateExit") && request.Parameters["FromDateExit"] != null && request.Parameters.ContainsKey("ToDateExit") && request.Parameters["ToDateExit"] != null)
                    {
                        string fromDateExit = request.Parameters["FromDateExit"]?.ToString();
                        string toDateExit = request.Parameters["ToDateExit"]?.ToString();
                        if (!string.IsNullOrEmpty(fromDateExit) && !string.IsNullOrEmpty(toDateExit)
                            && DateTime.TryParseExact(fromDateExit, "yyyy-MM-dd", CultureInfo.InvariantCulture, DateTimeStyles.None, out var parsedFromDateExit)
                            && DateTime.TryParseExact(toDateExit, "yyyy-MM-dd", CultureInfo.InvariantCulture, DateTimeStyles.None, out var parsedToDateExit))
                        {
                            if (filter.Length > 0) filter.Append(", ");
                            filter.Append($"מתאריך יציאה {parsedFromDateExit:dd/MM/yy} עד {parsedToDateExit:dd/MM/yy}");
                        }
                    }


                    worksheet.Cell(currentRow, 1).Value = filter.ToString();
                    worksheet.Range(worksheet.Cell(currentRow, 1), worksheet.Cell(currentRow, numberOfColumns)).Merge();
                    worksheet.Cell(currentRow, 1).Style.Font.Bold = true;
                    worksheet.Cell(currentRow, 1).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Right;
                    worksheet.Cell(currentRow, 1).Style.Alignment.WrapText = true;
                    worksheet.Row(currentRow).Height = 25;

                    currentRow += 2;

                    var table1 = worksheet.Cell(currentRow, 1).InsertTable(dataSet.Tables[0]);
                    ApplyTableStyleBoldHeadings(table1);


                    currentRow += dataSet.Tables[0].Rows.Count + 1;
                    worksheet.Cell(currentRow, 1).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Right;
                    worksheet.Range(currentRow, 1, currentRow, numberOfColumns).Style.Border.TopBorder = XLBorderStyleValues.Thin;

                    int startCalc = 6;
                    var columnsToCount = new List<string> { "מספר מכולה" };
                    ApplyColumnFormula(worksheet, dataSet.Tables[0], startCalc, currentRow, columnsToCount, "COUNTA");

                    int columnIndex = dataSet.Tables[0].Columns.IndexOf("מספר מכולה");
                    worksheet.Cell(currentRow, columnIndex).Value = ":מספר מכולות";
                    worksheet.Cell(currentRow, columnIndex).Style.Font.Bold = true;


                    ApplyNumberFormatToSheet(worksheet);

                    worksheet.Columns().AdjustToContents();

                    using (var stream = new MemoryStream())
                    {
                        workbook.SaveAs(stream);
                        return stream.ToArray();
                    }
                }
            }
            catch (Exception ex)
            {
                logger.WriteLog($"Error to Generate ContDtl Report: {ex}");
                throw new CustomException((int)ErrorMessages.ErrorCodes.GlobalError, ex.Message);
            }
        }

        private byte[] GeneratedeliveryCancel25Report(DataSet dataSet, ReportRequest request, ReportDtl reportDtl, Manifest manifest)
        {
            try
            {
                using (var workbook = new XLWorkbook())
                {
                    var worksheet = workbook.Worksheets.Add(reportDtl.ReportID);

                    PrintSettings(worksheet);

                    int currentRow = 1;
                    int numberOfColumns = dataSet.Tables[0].Columns.Count;

                    AddHeader2(worksheet, request, reportDtl, manifest, currentRow, numberOfColumns + 2);

                    currentRow += 2;

                    StringBuilder filter = new StringBuilder();

                    if (request.Parameters.ContainsKey("FromDate") && request.Parameters["FromDate"] != null && request.Parameters.ContainsKey("ToDate") && request.Parameters["ToDate"] != null)
                    {
                        string fromDate = request.Parameters["FromDate"]?.ToString();
                        string toDate = request.Parameters["ToDate"]?.ToString();
                        if (!string.IsNullOrEmpty(fromDate) && !string.IsNullOrEmpty(toDate)
                            && DateTime.TryParseExact(fromDate, "yyyy-MM-dd", CultureInfo.InvariantCulture, DateTimeStyles.None, out var parsedFromDate)
                            && DateTime.TryParseExact(toDate, "yyyy-MM-dd", CultureInfo.InvariantCulture, DateTimeStyles.None, out var parsedToDate))
                        {
                            filter.Append($"לתקופה: {parsedToDate:dd/MM/yy} - {parsedFromDate:dd/MM/yy}");
                        }
                    }

                    if (request.Parameters.ContainsKey("BilledImporterID") && request.Parameters["BilledImporterID"] != null)
                    {
                        string billedImporterID = request.Parameters["BilledImporterID"]?.ToString();
                        if (!string.IsNullOrEmpty(billedImporterID))
                        {
                            string billedImporter = dataSet.Tables[1].Rows[0]["BilledImporter"].ToString();
                            if (filter.Length > 0) filter.Append(", ");
                            filter.Append($"ללקוח: {billedImporter} ({billedImporterID})");
                        }
                    }

                    if (request.Parameters.ContainsKey("FromGush") && request.Parameters["FromGush"] != null && request.Parameters.ContainsKey("ToGush") && request.Parameters["ToGush"] != null)
                    {
                        string fromGush = request.Parameters["FromGush"]?.ToString();
                        string toGush = request.Parameters["ToGush"]?.ToString();
                        if (!string.IsNullOrEmpty(fromGush) && !string.IsNullOrEmpty(toGush))
                        {
                            string FormattedFromGush = fromGush.Length > 2 ? $"{fromGush.Substring(0, 2)}/{fromGush.Substring(2)}" : fromGush;
                            string FormattedToGush = toGush.Length > 2 ? $"{toGush.Substring(0, 2)}/{toGush.Substring(2)}" : toGush;

                            if (filter.Length > 0) filter.Append(", ");
                            filter.Append($"גושים {FormattedToGush} - {FormattedFromGush}");
                        }
                    }


                    worksheet.Cell(currentRow, 1).Value = filter.ToString();
                    worksheet.Range(worksheet.Cell(currentRow, 1), worksheet.Cell(currentRow, numberOfColumns)).Merge();
                    worksheet.Cell(currentRow, 1).Style.Font.Bold = true;
                    worksheet.Cell(currentRow, 1).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Right;

                    currentRow += 2;

                    var table1 = worksheet.Cell(currentRow, 1).InsertTable(dataSet.Tables[0]);
                    ApplyTableStyleBoldHeadings(table1);


                    currentRow += dataSet.Tables[0].Rows.Count + 1;
                    worksheet.Cell(currentRow, 1).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Right;
                    worksheet.Range(currentRow, 1, currentRow, numberOfColumns).Style.Border.TopBorder = XLBorderStyleValues.Thin;


                    ApplyNumberFormatToSheet(worksheet);

                    worksheet.Columns().AdjustToContents();


                    using (var stream = new MemoryStream())
                    {
                        workbook.SaveAs(stream);
                        return stream.ToArray();
                    }
                }
            }
            catch (Exception ex)
            {
                logger.WriteLog($"Error to Generate deliveryCancel25 Report: {ex}");
                throw new CustomException((int)ErrorMessages.ErrorCodes.GlobalError, ex.Message);
            }
        }

        private byte[] GenerateRelease33Report(DataSet dataSet, ReportRequest request, ReportDtl reportDtl, Manifest manifest)
        {
            try
            {
                using (var workbook = new XLWorkbook())
                {
                    var worksheet = workbook.Worksheets.Add(reportDtl.ReportID);

                    PrintSettings(worksheet);

                    int currentRow = 1;
                    int numberOfColumns = dataSet.Tables[0].Columns.Count;

                    AddHeader2(worksheet, request, reportDtl, manifest, currentRow, numberOfColumns);

                    currentRow += 2;

                    StringBuilder filter = new StringBuilder();

                    if (request.Parameters.ContainsKey("FromDate") && request.Parameters["FromDate"] != null && request.Parameters.ContainsKey("ToDate") && request.Parameters["ToDate"] != null)
                    {
                        string fromDate = request.Parameters["FromDate"]?.ToString();
                        string toDate = request.Parameters["ToDate"]?.ToString();
                        if (!string.IsNullOrEmpty(fromDate) && !string.IsNullOrEmpty(toDate)
                            && DateTime.TryParseExact(fromDate, "yyyy-MM-dd", CultureInfo.InvariantCulture, DateTimeStyles.None, out var parsedFromDate)
                            && DateTime.TryParseExact(toDate, "yyyy-MM-dd", CultureInfo.InvariantCulture, DateTimeStyles.None, out var parsedToDate))
                        {
                            filter.Append($"לתקופה: מ {parsedFromDate:dd/MM/yy} עד {parsedToDate:dd/MM/yy}");
                        }
                    }

                    if (request.Parameters.ContainsKey("FromGush") && request.Parameters["FromGush"] != null && request.Parameters.ContainsKey("ToGush") && request.Parameters["ToGush"] != null)
                    {
                        string fromGush = request.Parameters["FromGush"]?.ToString();
                        string toGush = request.Parameters["ToGush"]?.ToString();
                        if (!string.IsNullOrEmpty(fromGush) && !string.IsNullOrEmpty(toGush))
                        {
                            string FormattedFromGush = fromGush.Length > 2 ? $"{fromGush.Substring(0, 2)}/{fromGush.Substring(2)}" : fromGush;
                            string FormattedToGush = toGush.Length > 2 ? $"{toGush.Substring(0, 2)}/{toGush.Substring(2)}" : toGush;

                            if (filter.Length > 0) filter.Append(", ");
                            filter.Append($"לגושים {FormattedFromGush} עד {FormattedToGush}");
                        }
                    }

                    if (request.Parameters.ContainsKey("BilledImporterID") && request.Parameters["BilledImporterID"] != null)
                    {
                        string billedImporterID = request.Parameters["BilledImporterID"]?.ToString();
                        if (!string.IsNullOrEmpty(billedImporterID))
                        {
                            string billedImporter = dataSet.Tables[1].Rows[0]["BilledImporter"].ToString();
                            if (filter.Length > 0) filter.Append(", ");
                            filter.Append($"ללקוח: {billedImporter} ({billedImporterID})");
                        }
                    }


                    worksheet.Cell(currentRow, 1).Value = filter.ToString();
                    worksheet.Range(worksheet.Cell(currentRow, 1), worksheet.Cell(currentRow, numberOfColumns)).Merge();
                    worksheet.Cell(currentRow, 1).Style.Font.Bold = true;
                    worksheet.Cell(currentRow, 1).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Right;

                    currentRow += 2;

                    var columnsToSum = new List<string> { "כמות בהתרה", "ערך בשקלים" };
                    var columnsToCount = new List<string> { "התרת שחרור" };

                    if (request.IsPrint)
                    {
                        string groupColumnName = "שם לקוח";

                        InsertTableWithSubtotals(worksheet, dataSet.Tables[0], groupColumnName, columnsToSum, columnsToCount, currentRow);
                    }

                    else
                    {
                        var table1 = worksheet.Cell(currentRow, 1).InsertTable(dataSet.Tables[0]);
                        ApplyTableStyleBoldHeadings(table1);

                        currentRow += dataSet.Tables[0].Rows.Count + 1;
                        worksheet.Cell(currentRow, 1).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Right;
                        worksheet.Range(currentRow, 1, currentRow, numberOfColumns).Style.Border.TopBorder = XLBorderStyleValues.Thin;

                        int startCalc = 6;
                        ApplyColumnFormula(worksheet, dataSet.Tables[0], startCalc, currentRow, columnsToSum, "SUM");
                        ApplyColumnFormula(worksheet, dataSet.Tables[0], startCalc, currentRow, columnsToCount, "COUNTA");
                    }

                    ApplyNumberFormatToSheet(worksheet);

                    worksheet.Columns().AdjustToContents();

                    using (var stream = new MemoryStream())
                    {
                        workbook.SaveAs(stream);
                        return stream.ToArray();
                    }
                }
            }
            catch (Exception ex)
            {
                logger.WriteLog($"Error to Generate Release33 Report: {ex}");
                throw new CustomException((int)ErrorMessages.ErrorCodes.GlobalError, ex.Message);
            }
        }

        private byte[] GenerateReleaseGoods10Report(DataSet dataSet, ReportRequest request, ReportDtl reportDtl, Manifest manifest)
        {
            try
            {
                using (var workbook = new XLWorkbook())
                {
                    var worksheet = workbook.Worksheets.Add(reportDtl.ReportID);

                    PrintSettings(worksheet);

                    int currentRow = 1;
                    int numberOfColumns = dataSet.Tables[0].Columns.Count;

                    AddHeader2(worksheet, request, reportDtl, manifest, currentRow, numberOfColumns);

                    currentRow += 2;

                    // Filter-summary line. Column order/headers for the data itself come entirely
                    // from the stored procedure's SELECT list (Tables[0]) - nothing here decides
                    // column layout, only the free-text description of which filters were applied.
                    StringBuilder filter = new StringBuilder();

                    if (request.Parameters.ContainsKey("ValidityDate") && request.Parameters["ValidityDate"] != null)
                    {
                        string validityDate = request.Parameters["ValidityDate"]?.ToString();
                        if (!string.IsNullOrEmpty(validityDate)
                            && DateTime.TryParseExact(validityDate, "yyyy-MM-dd", CultureInfo.InvariantCulture, DateTimeStyles.None, out var parsedValidityDate))
                        {
                            filter.Append($"נכון לתאריך: {parsedValidityDate:dd/MM/yy}");
                        }
                    }

                    if (request.Parameters.ContainsKey("BilledImporterID") && request.Parameters["BilledImporterID"] != null)
                    {
                        string billedImporterID = request.Parameters["BilledImporterID"]?.ToString();
                        if (!string.IsNullOrEmpty(billedImporterID))
                        {
                            string billedImporter = dataSet.Tables[1].Rows.Count > 0 ? dataSet.Tables[1].Rows[0]["BilledImporter"].ToString() : billedImporterID;
                            if (filter.Length > 0) filter.Append(", ");
                            filter.Append($"ללקוח: {billedImporter} ({billedImporterID})");
                        }
                    }

                    if (request.Parameters.ContainsKey("FromGush") && request.Parameters["FromGush"] != null && request.Parameters.ContainsKey("ToGush") && request.Parameters["ToGush"] != null)
                    {
                        string fromGush = request.Parameters["FromGush"]?.ToString();
                        string toGush = request.Parameters["ToGush"]?.ToString();
                        if (!string.IsNullOrEmpty(fromGush) && !string.IsNullOrEmpty(toGush))
                        {
                            string formattedFromGush = fromGush.Length > 2 ? $"{fromGush.Substring(0, 2)}/{fromGush.Substring(2)}" : fromGush;
                            string formattedToGush = toGush.Length > 2 ? $"{toGush.Substring(0, 2)}/{toGush.Substring(2)}" : toGush;

                            if (filter.Length > 0) filter.Append(", ");
                            filter.Append($"לגושים {formattedFromGush} עד {formattedToGush}");
                        }
                    }

                    // "התרת שחרור" label shown whenever either a specific declaration or a date
                    // range was supplied - per spec note: show the label even without a decleration
                    // ID, as long as a period was given.
                    bool hasDeclerationID = request.Parameters.ContainsKey("declerationID") && request.Parameters["declerationID"] != null
                        && !string.IsNullOrEmpty(request.Parameters["declerationID"]?.ToString());
                    bool hasDateRange = request.Parameters.ContainsKey("FromDate") && request.Parameters["FromDate"] != null
                        && request.Parameters.ContainsKey("ToDate") && request.Parameters["ToDate"] != null;

                    if (hasDeclerationID || hasDateRange)
                    {
                        if (filter.Length > 0) filter.Append(", ");
                        filter.Append("התרת שחרור");

                        if (hasDeclerationID)
                        {
                            filter.Append($" {request.Parameters["declerationID"]}");
                        }

                        if (hasDateRange)
                        {
                            string fromDate = request.Parameters["FromDate"]?.ToString();
                            string toDate = request.Parameters["ToDate"]?.ToString();
                            if (DateTime.TryParseExact(fromDate, "yyyy-MM-dd", CultureInfo.InvariantCulture, DateTimeStyles.None, out var parsedFromDate)
                                && DateTime.TryParseExact(toDate, "yyyy-MM-dd", CultureInfo.InvariantCulture, DateTimeStyles.None, out var parsedToDate))
                            {
                                filter.Append($" מתאריך {parsedFromDate:dd/MM/yy} עד תאריך {parsedToDate:dd/MM/yy}");
                            }
                        }
                    }

                    worksheet.Cell(currentRow, 1).Value = filter.ToString();
                    worksheet.Range(worksheet.Cell(currentRow, 1), worksheet.Cell(currentRow, numberOfColumns)).Merge();
                    worksheet.Cell(currentRow, 1).Style.Font.Bold = true;
                    worksheet.Cell(currentRow, 1).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Right;

                    currentRow += 2;

                    // Columns whose header text is fixed by the stored procedure - only listed
                    // here by name so the subtotal/total rows know which cells to sum.
                    var columnsToSum = new List<string> { "נקלט", "שוחרר", "נמסר", "יתרה משוחררת", "ברשות מוסמכת" };
                    const string groupColumnName = "גוש";
                    const string labelColumnName = "שם לקוח";

                    // NOTE (flagged for review): the "received" and "delivered" columns are, per
                    // spec, a single value per gush / per declaration respectively, repeated on
                    // every line row of that group (per explicit instruction). A plain SUM() over
                    // those repeated values would over-count whenever a group has more than one
                    // line. The "released" and "authority" columns are genuine per-line values
                    // and sum correctly as-is. To keep the displayed subtotal numerically correct
                    // while still following "repeat + sum" for the on-screen column, deduping is
                    // applied for the two repeated columns specifically when building the
                    // subtotal/total rows below.

                    DataTable source = dataSet.Tables[0];
                    DataTable output = source.Clone();
                    output.Columns.Add("IsSummaryRow", typeof(bool));

                    var groups = source.AsEnumerable().GroupBy(r => r[groupColumnName]);

                    foreach (var group in groups)
                    {
                        var groupRows = group.ToList();

                        foreach (var row in groupRows)
                        {
                            var newRow = output.NewRow();
                            newRow.ItemArray = row.ItemArray;
                            newRow["IsSummaryRow"] = false;
                            output.Rows.Add(newRow);
                        }

                        var summaryRow = output.NewRow();
                        summaryRow[labelColumnName] = "סה\"כ התרות";

                        summaryRow["שוחרר"] = groupRows.Sum(r => decimal.TryParse(r["שוחרר"]?.ToString(), out var v) ? v : 0);
                        summaryRow["ברשות מוסמכת"] = groupRows.Sum(r => decimal.TryParse(r["ברשות מוסמכת"]?.ToString(), out var v) ? v : 0);

                        // Received quantity: constant per gush - take it once rather than summing repeats.
                        summaryRow["נקלט"] = groupRows
                            .Select(r => decimal.TryParse(r["נקלט"]?.ToString(), out var v) ? v : 0)
                            .FirstOrDefault();

                        // Delivered quantity: constant per release (DeclarationID) - dedupe by
                        // declaration before summing, so a multi-line declaration isn't counted
                        // once per line.
                        decimal delivered = groupRows
                            .GroupBy(r => r["התרת שחרור"]?.ToString())
                            .Select(g => decimal.TryParse(g.First()["נמסר"]?.ToString(), out var v) ? v : 0)
                            .Sum();
                        summaryRow["נמסר"] = delivered;

                        summaryRow["יתרה משוחררת"] = Convert.ToDecimal(summaryRow["שוחרר"]) - delivered;

                        summaryRow["IsSummaryRow"] = true;
                        output.Rows.Add(summaryRow);
                    }

                    var totalRow = output.NewRow();
                    totalRow[labelColumnName] = "סה\"כ";

                    totalRow["שוחרר"] = source.AsEnumerable().Sum(r => decimal.TryParse(r["שוחרר"]?.ToString(), out var v) ? v : 0);
                    totalRow["ברשות מוסמכת"] = source.AsEnumerable().Sum(r => decimal.TryParse(r["ברשות מוסמכת"]?.ToString(), out var v) ? v : 0);

                    totalRow["נקלט"] = source.AsEnumerable()
                        .GroupBy(r => r[groupColumnName])
                        .Select(g => decimal.TryParse(g.First()["נקלט"]?.ToString(), out var v) ? v : 0)
                        .Sum();

                    decimal totalDelivered = source.AsEnumerable()
                        .GroupBy(r => r["התרת שחרור"]?.ToString())
                        .Select(g => decimal.TryParse(g.First()["נמסר"]?.ToString(), out var v) ? v : 0)
                        .Sum();
                    totalRow["נמסר"] = totalDelivered;

                    totalRow["יתרה משוחררת"] = Convert.ToDecimal(totalRow["שוחרר"]) - totalDelivered;

                    totalRow["IsSummaryRow"] = true;
                    output.Rows.Add(totalRow);

                    var tableRange = worksheet.Cell(currentRow, 1).InsertTable(output);
                    ApplyTableStyleBoldHeadings(tableRange);

                    int summaryColIndex = output.Columns["IsSummaryRow"].Ordinal + 1;

                    foreach (var row in tableRange.Rows())
                    {
                        bool.TryParse(row.Cell(summaryColIndex).GetValue<string>(), out bool isSummary);
                        if (isSummary)
                        {
                            row.Style.Font.Bold = true;
                            row.Style.Border.BottomBorder = XLBorderStyleValues.Thin;
                        }
                    }

                    worksheet.Column(summaryColIndex).Delete();

                    ApplyNumberFormatToSheet(worksheet);

                    worksheet.Columns().AdjustToContents();

                    using (var stream = new MemoryStream())
                    {
                        workbook.SaveAs(stream);
                        return stream.ToArray();
                    }
                }
            }
            catch (Exception ex)
            {
                logger.WriteLog($"Error to Generate ReleaseGoods10 Report: {ex}");
                throw new CustomException((int)ErrorMessages.ErrorCodes.GlobalError, ex.Message);
            }
        }

        private byte[] GenerateCustomerInvCover28Report(DataSet dataSet, ReportRequest request, ReportDtl reportDtl, Manifest manifest)
        {
            try
            {
                using (var workbook = new XLWorkbook())
                {
                    var worksheet = workbook.Worksheets.Add(reportDtl.ReportID);

                    PrintSettings(worksheet);

                    int currentRow = 1;
                    int numberOfColumns = dataSet.Tables[0].Columns.Count;

                    AddHeader2(worksheet, request, reportDtl, manifest, currentRow, numberOfColumns + 6);

                    currentRow += 2;

                    StringBuilder filter = new StringBuilder();

                    if (request.Parameters.ContainsKey("BilledImporterID") && request.Parameters["BilledImporterID"] != null)
                    {
                        string billedImporterID = request.Parameters["BilledImporterID"]?.ToString();
                        if (!string.IsNullOrEmpty(billedImporterID))
                        {
                            string billedImporter = dataSet.Tables[1].Rows[0]["BilledImporter"].ToString();
                            if (filter.Length > 0) filter.Append(", ");
                            filter.Append($"ללקוח: {billedImporter} ({billedImporterID})");
                        }
                    }


                    worksheet.Cell(currentRow, 1).Value = filter.ToString();
                    worksheet.Range(worksheet.Cell(currentRow, 1), worksheet.Cell(currentRow, numberOfColumns + 6)).Merge();
                    worksheet.Cell(currentRow, 1).Style.Font.Bold = true;
                    worksheet.Cell(currentRow, 1).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;

                    currentRow += 2;

                    var table1 = worksheet.Cell(currentRow, 1).InsertTable(dataSet.Tables[0]);
                    ApplyTableStyleBoldHeadings(table1);


                    currentRow += dataSet.Tables[0].Rows.Count + 1;
                    worksheet.Cell(currentRow, 1).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Right;
                    worksheet.Range(currentRow, 1, currentRow, numberOfColumns).Style.Border.TopBorder = XLBorderStyleValues.Thin;


                    var rows = table1.DataRange.Rows().ToList();

                    int dateColumnIndex = dataSet.Tables[0].Columns["תוקף כיסוי נפח"].Ordinal + 1;

                    string lastDate = null;

                    foreach (var row in rows)
                    {
                        var currDate = row.Cell(dateColumnIndex).GetString();

                        if (lastDate != null && currDate != lastDate)
                            row.Style.Border.TopBorder = XLBorderStyleValues.Thin;

                        if (!string.IsNullOrEmpty(currDate))
                            lastDate = currDate;
                    }


                    ApplyNumberFormatToSheet(worksheet);

                    worksheet.Columns().AdjustToContents();

                    using (var stream = new MemoryStream())
                    {
                        workbook.SaveAs(stream);
                        return stream.ToArray();
                    }
                }
            }
            catch (Exception ex)
            {
                logger.WriteLog($"Error to Generate CustomerInvCover28 Report: {ex}");
                throw new CustomException((int)ErrorMessages.ErrorCodes.GlobalError, ex.Message);
            }
        }

        private byte[] GenerateStorageCalcReport(DataSet dataSet, ReportRequest request, ReportDtl reportDtl, Manifest manifest)
        {
            try
            {
                using (var workbook = new XLWorkbook())
                {
                    var worksheet = workbook.Worksheets.Add(reportDtl.ReportID);

                    if (request.Parameters.ContainsKey("ImporterName") && request.Parameters["ImporterName"] != null)
                    {
                        string importerName = request.Parameters["ImporterName"]?.ToString();
                        if (!string.IsNullOrEmpty(importerName))
                        {
                            var columnsToDelete = new List<string> { "לקוח" };
                            RemoveColumnsByName(dataSet.Tables[0], columnsToDelete);
                        }
                    }

                    PrintSettings(worksheet);

                    int currentRow = 1;
                    int numberOfColumns = dataSet.Tables[0].Columns.Count;

                    currentRow++;

                    worksheet.Cell(currentRow, 1).Value = $"{manifest.TermName}, {manifest.TermVAT}";
                    worksheet.Cell(currentRow, 1).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Right;
                    worksheet.Cell(currentRow, 1).Style.Font.Bold = true;
                    worksheet.Cell(currentRow, 1).Style.Font.FontSize = 11;
                    worksheet.Range(worksheet.Cell(currentRow, 1), worksheet.Cell(currentRow, 4)).Merge();

                    string base64String = manifest.Logo.Split(',')[1];
                    byte[] imageBytes = Convert.FromBase64String(base64String);
                    MemoryStream imageStream = new MemoryStream(imageBytes);

                    int startColumn = numberOfColumns - 1;
                    int columnsToSpan = 2;

                    // Calculating the total width of the two columns
                    double totalWidthUnits = 0;
                    for (int i = 0; i < columnsToSpan; i++)
                    {
                        totalWidthUnits += worksheet.Column(startColumn + i).Width;
                    }

                    // Converting from Excel width units to pixels – estimate: approximately 7.5 pixels per width unit
                    int widthInPixels = (int)(totalWidthUnits * 7.5);

                    var picture = worksheet.AddPicture(imageStream).MoveTo(worksheet.Cell(1, startColumn)).WithSize(widthInPixels, 70);

                    currentRow++;

                    worksheet.Cell(currentRow, 1).Value = $"{manifest.Location}, ת.ד. {manifest.POB}, מיקוד {manifest.ZIPcode}";
                    worksheet.Cell(currentRow, 1).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Right;
                    worksheet.Cell(currentRow, 1).Style.Font.Bold = true;
                    worksheet.Cell(currentRow, 1).Style.Font.FontSize = 11;
                    worksheet.Range(worksheet.Cell(currentRow, 1), worksheet.Cell(currentRow, 4)).Merge();

                    currentRow++;

                    worksheet.Cell(currentRow, 1).Value = $"{manifest.Phone}";
                    worksheet.Cell(currentRow, 1).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Right;
                    worksheet.Cell(currentRow, 1).Style.Font.Bold = true;
                    worksheet.Cell(currentRow, 1).Style.Font.FontSize = 11;
                    worksheet.Range(worksheet.Cell(currentRow, 1), worksheet.Cell(currentRow, 4)).Merge();

                    currentRow++;

                    worksheet.Range(currentRow, 1, currentRow, numberOfColumns).Style.Border.TopBorder = XLBorderStyleValues.Thin;


                    StringBuilder header = new StringBuilder();
                    header.Append(reportDtl.ReportName);

                    if (request.Parameters.ContainsKey("direction") && request.Parameters["direction"] != null)
                    {
                        string direction = request.Parameters["direction"]?.ToString();
                        if (!string.IsNullOrEmpty(direction))
                        {
                            if (direction == "I")
                                header.Append(" ליבוא");
                            else
                                header.Append(" ליצוא");

                        }
                    }

                    if (request.Parameters.ContainsKey("ImporterName") && request.Parameters["ImporterName"] != null)
                    {
                        string importerName = request.Parameters["ImporterName"]?.ToString();
                        if (!string.IsNullOrEmpty(importerName))
                        {
                            header.Append($", עבור {importerName}");

                        }
                    }

                    worksheet.Cell(currentRow, 1).Value = header.ToString();
                    worksheet.Cell(currentRow, 1).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                    worksheet.Cell(currentRow, 1).Style.Font.Bold = true;
                    worksheet.Cell(currentRow, 1).Style.Font.FontSize = 11;
                    worksheet.Range(worksheet.Cell(currentRow, 1), worksheet.Cell(currentRow, numberOfColumns)).Merge();

                    currentRow++;

                    worksheet.Cell(currentRow, 1).Value = $"נכון ל - {DateTime.Now.ToString("dd/MM/yy")}";
                    worksheet.Cell(currentRow, 1).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                    worksheet.Cell(currentRow, 1).Style.Font.Bold = true;
                    worksheet.Range(worksheet.Cell(currentRow, 1), worksheet.Cell(currentRow, numberOfColumns)).Merge();

                    currentRow += 4;

                    var table1 = worksheet.Cell(currentRow, 1).InsertTable(dataSet.Tables[0]);
                    ApplyTableStyleBoldHeadings(table1);

                    currentRow += dataSet.Tables[0].Rows.Count + 1;
                    worksheet.Cell(currentRow, 1).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Right;
                    worksheet.Range(currentRow, 1, currentRow, numberOfColumns).Style.Border.TopBorder = XLBorderStyleValues.Thin;

                    int startCalc = 11;

                    int columnIndex = dataSet.Tables[0].Columns["סה\"כ בש\"ח"].Ordinal + 1;

                    decimal total = 0;

                    for (int row = startCalc; row < currentRow; row++)
                    {
                        var cell = worksheet.Cell(row, columnIndex);
                        if (decimal.TryParse(cell.GetValue<string>(), out decimal value))
                        {
                            total += value;
                        }
                    }

                    worksheet.Cell(currentRow, columnIndex - 1).Value = "סכום כולל:";
                    worksheet.Cell(currentRow, columnIndex - 1).Style.Font.Bold = true;

                    worksheet.Cell(currentRow, columnIndex).Value = $"₪ {total.ToString("N0")}";
                    worksheet.Cell(currentRow, columnIndex).Style.Font.Bold = true;
                    worksheet.Cell(currentRow, columnIndex).Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Right);

                    currentRow++;

                    decimal vat = Convert.ToDecimal(request.Parameters["VAT"]);
                    decimal vatAmount = total * vat;

                    worksheet.Cell(currentRow, columnIndex - 1).Value = $"מע\"מ ({(int)(vat * 100)}%):";
                    worksheet.Cell(currentRow, columnIndex - 1).Style.Font.Bold = true;

                    worksheet.Cell(currentRow, columnIndex).Value = $"₪ {vatAmount.ToString("N0")}";
                    worksheet.Cell(currentRow, columnIndex).Style.Font.Bold = true;
                    worksheet.Cell(currentRow, columnIndex).Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Right);

                    currentRow++;

                    worksheet.Cell(currentRow, columnIndex - 1).Value = "סה\"כ:";
                    worksheet.Cell(currentRow, columnIndex - 1).Style.Font.Bold = true;

                    worksheet.Cell(currentRow, columnIndex).Value = $"₪ {(total + vatAmount).ToString("N0")}";
                    worksheet.Cell(currentRow, columnIndex).Style.Font.Bold = true;
                    worksheet.Cell(currentRow, columnIndex).Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Right);


                    int lastRow = worksheet.LastRowUsed().RowNumber();

                    worksheet.Range(startCalc, 1, lastRow, 1).Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Left);

                    ApplyNumberFormatToSheet(worksheet);

                    worksheet.Columns().AdjustToContents();

                    using (var stream = new MemoryStream())
                    {
                        workbook.SaveAs(stream);
                        return stream.ToArray();
                    }
                }
            }
            catch (Exception ex)
            {
                logger.WriteLog($"Error to Generate StorageCalc Report: {ex}");
                throw new CustomException((int)ErrorMessages.ErrorCodes.GlobalError, ex.Message);
            }
        }

        private byte[] GenerateExtAutorityInv18Report(DataSet dataSet, ReportRequest request, ReportDtl reportDtl, Manifest manifest)
        {
            try
            {
                using (var workbook = new XLWorkbook())
                {
                    var worksheet = workbook.Worksheets.Add(reportDtl.ReportID);

                    PrintSettings(worksheet);

                    int currentRow = 1;
                    int numberOfColumns = dataSet.Tables[0].Columns.Count;

                    AddHeader2(worksheet, request, reportDtl, manifest, currentRow, numberOfColumns);

                    currentRow += 2;

                    StringBuilder filter = new StringBuilder();

                    if (request.Parameters.ContainsKey("BilledImporterID") && request.Parameters["BilledImporterID"] != null)
                    {
                        string billedImporterID = request.Parameters["BilledImporterID"]?.ToString();
                        if (!string.IsNullOrEmpty(billedImporterID))
                        {
                            string billedImporter = dataSet.Tables[1].Rows[0]["BilledImporter"].ToString();
                            if (filter.Length > 0) filter.Append(", ");
                            filter.Append($"עבור לקוח: {billedImporter} ({billedImporterID})");
                        }
                    }

                    if (request.Parameters.ContainsKey("FromGush") && request.Parameters["FromGush"] != null && request.Parameters.ContainsKey("ToGush") && request.Parameters["ToGush"] != null)
                    {
                        string fromGush = request.Parameters["FromGush"]?.ToString();
                        string toGush = request.Parameters["ToGush"]?.ToString();
                        if (!string.IsNullOrEmpty(fromGush) && !string.IsNullOrEmpty(toGush))
                        {
                            string FormattedFromGush = fromGush.Length > 2 ? $"{fromGush.Substring(0, 2)}/{fromGush.Substring(2)}" : fromGush;
                            string FormattedToGush = toGush.Length > 2 ? $"{toGush.Substring(0, 2)}/{toGush.Substring(2)}" : toGush;

                            if (filter.Length > 0) filter.Append(", ");
                            filter.Append($"מגוש {FormattedFromGush} עד גוש {FormattedToGush}");
                        }
                    }

                    if (request.Parameters.ContainsKey("AuthorityCode") && request.Parameters["AuthorityCode"] != null)
                    {
                        string authorityCode = request.Parameters["AuthorityCode"]?.ToString();
                        if (!string.IsNullOrEmpty(authorityCode))
                        {
                            string description = dataSet.Tables[2].Rows[0]["Description"].ToString();
                            if (filter.Length > 0) filter.Append(", ");
                            filter.Append($"רשות: {description}");
                        }
                    }

                    if (request.Parameters.ContainsKey("AlreadyClosed") && request.Parameters["AlreadyClosed"] != null)
                    {
                        string alreadyClosed = request.Parameters["AlreadyClosed"]?.ToString();
                        if (!string.IsNullOrEmpty(alreadyClosed) && alreadyClosed == "Y")
                        {
                            if (filter.Length > 0) filter.Append(", ");
                            filter.Append("כולל גושים סגורים");
                        }
                    }

                    if (request.Parameters.ContainsKey("ZeroInventory") && request.Parameters["ZeroInventory"] != null)
                    {
                        string zeroInventory = request.Parameters["ZeroInventory"]?.ToString();
                        if (!string.IsNullOrEmpty(zeroInventory) && zeroInventory == "Y")
                        {
                            if (filter.Length > 0) filter.Append(", ");
                            filter.Append("כולל יתרות אפס ברשות");
                        }
                    }


                    worksheet.Cell(currentRow, 1).Value = filter.ToString();
                    worksheet.Range(worksheet.Cell(currentRow, 1), worksheet.Cell(currentRow, numberOfColumns)).Merge();
                    worksheet.Cell(currentRow, 1).Style.Font.Bold = true;
                    worksheet.Cell(currentRow, 1).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Right;

                    currentRow += 2;

                    var columnsToSum = new List<string> { "יתרה ברשות" };
                    var columnsToCount = new List<string> { };

                    string groupColumnName = "גוש";

                    InsertTableWithSubtotals(worksheet, dataSet.Tables[0], groupColumnName, columnsToSum, columnsToCount, currentRow);

                    int totalRows = dataSet.Tables[0].Rows.Count + dataSet.Tables[0].AsEnumerable().GroupBy(r => r[groupColumnName]).Count() + currentRow + 1;

                    worksheet.Cell(totalRows, numberOfColumns - 1).Value = "יתרה לדוח:";

                    ApplyNumberFormatToSheet(worksheet);

                    worksheet.Columns().AdjustToContents();


                    using (var stream = new MemoryStream())
                    {
                        workbook.SaveAs(stream);
                        return stream.ToArray();
                    }
                }
            }
            catch (Exception ex)
            {
                logger.WriteLog($"Error to Generate ExtAutorityInv18 Report: {ex}");
                throw new CustomException((int)ErrorMessages.ErrorCodes.GlobalError, ex.Message);
            }
        }

        private byte[] GenerateCustomerInvAvg4Report(DataSet dataSet, ReportRequest request, ReportDtl reportDtl, Manifest manifest)
        {
            try
            {
                using (var workbook = new XLWorkbook())
                {
                    var worksheet = workbook.Worksheets.Add(reportDtl.ReportID);

                    PrintSettings(worksheet);

                    int currentRow = 1;
                    int numberOfColumns = dataSet.Tables[0].Columns.Count;

                    AddHeader2(worksheet, request, reportDtl, manifest, currentRow, numberOfColumns);

                    currentRow += 2;

                    StringBuilder filter = new StringBuilder();

                    if (request.Parameters.ContainsKey("BilledImporterID") && request.Parameters["BilledImporterID"] != null)
                    {
                        string billedImporterID = request.Parameters["BilledImporterID"]?.ToString();
                        if (!string.IsNullOrEmpty(billedImporterID))
                        {
                            string billedImporter = dataSet.Tables[1].Rows[0]["BilledImporter"].ToString();
                            if (filter.Length > 0) filter.Append(", ");
                            filter.Append($"ללקוח: {billedImporter} ({billedImporterID})");
                        }
                    }


                    worksheet.Cell(currentRow, 1).Value = filter.ToString();
                    worksheet.Range(worksheet.Cell(currentRow, 1), worksheet.Cell(currentRow, numberOfColumns)).Merge();
                    worksheet.Cell(currentRow, 1).Style.Font.Bold = true;
                    worksheet.Cell(currentRow, 1).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;

                    currentRow += 2;

                    var table1 = worksheet.Cell(currentRow, 1).InsertTable(dataSet.Tables[0]);
                    ApplyTableStyleBoldHeadings(table1);


                    currentRow += dataSet.Tables[0].Rows.Count + 1;
                    worksheet.Cell(currentRow, 1).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Right;
                    worksheet.Range(currentRow, 1, currentRow, numberOfColumns).Style.Border.TopBorder = XLBorderStyleValues.Thin;


                    ApplyNumberFormatToSheet(worksheet);

                    worksheet.Columns().AdjustToContents();


                    using (var stream = new MemoryStream())
                    {
                        workbook.SaveAs(stream);
                        return stream.ToArray();
                    }
                }
            }
            catch (Exception ex)
            {
                logger.WriteLog($"Error to Generate CustomerInvAvg4 Report: {ex}");
                throw new CustomException((int)ErrorMessages.ErrorCodes.GlobalError, ex.Message);
            }
        }

        private byte[] GenerateAVGStorageDays35Report(DataSet dataSet, ReportRequest request, ReportDtl reportDtl, Manifest manifest)
        {
            try
            {
                using (var workbook = new XLWorkbook())
                {
                    var worksheet = workbook.Worksheets.Add(reportDtl.ReportID);

                    PrintSettings(worksheet);

                    int currentRow = 1;
                    int numberOfColumns = dataSet.Tables[0].Columns.Count;

                    AddHeader2(worksheet, request, reportDtl, manifest, currentRow, numberOfColumns + 4);

                    currentRow += 2;

                    StringBuilder filter = new StringBuilder();

                    if (request.Parameters.ContainsKey("BilledImporterID") && request.Parameters["BilledImporterID"] != null)
                    {
                        string billedImporterID = request.Parameters["BilledImporterID"]?.ToString();
                        if (!string.IsNullOrEmpty(billedImporterID))
                        {
                            string billedImporter = dataSet.Tables[1].Rows[0]["BilledImporter"].ToString();
                            if (filter.Length > 0) filter.Append(", ");
                            filter.Append($"ללקוח: {billedImporter} ({billedImporterID})");
                        }
                    }

                    if (request.Parameters.ContainsKey("FromDate") && request.Parameters["FromDate"] != null && request.Parameters.ContainsKey("ToDate") && request.Parameters["ToDate"] != null)
                    {
                        string fromDate = request.Parameters["FromDate"]?.ToString();
                        string toDate = request.Parameters["ToDate"]?.ToString();
                        if (!string.IsNullOrEmpty(fromDate) && !string.IsNullOrEmpty(toDate)
                            && DateTime.TryParseExact(fromDate, "yyyy-MM-dd", CultureInfo.InvariantCulture, DateTimeStyles.None, out var parsedFromDate)
                            && DateTime.TryParseExact(toDate, "yyyy-MM-dd", CultureInfo.InvariantCulture, DateTimeStyles.None, out var parsedToDate))
                        {
                            filter.Append(Environment.NewLine);
                            filter.Append($"לתקופה: {parsedToDate:dd/MM/yy} - {parsedFromDate:dd/MM/yy}");
                        }
                    }


                    worksheet.Cell(currentRow, 1).Value = filter.ToString();
                    worksheet.Range(worksheet.Cell(currentRow, 1), worksheet.Cell(currentRow, numberOfColumns + 4)).Merge();
                    worksheet.Cell(currentRow, 1).Style.Font.Bold = true;
                    worksheet.Cell(currentRow, 1).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                    worksheet.Row(currentRow).Height = 25;

                    currentRow += 2;

                    var table1 = worksheet.Cell(currentRow, 1).InsertTable(dataSet.Tables[0]);
                    ApplyTableStyleBoldHeadings(table1);


                    currentRow += dataSet.Tables[0].Rows.Count + 1;
                    worksheet.Cell(currentRow, 1).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Right;
                    worksheet.Range(currentRow, 1, currentRow, numberOfColumns).Style.Border.TopBorder = XLBorderStyleValues.Thin;


                    ApplyNumberFormatToSheet(worksheet);

                    worksheet.Columns().AdjustToContents();


                    using (var stream = new MemoryStream())
                    {
                        workbook.SaveAs(stream);
                        return stream.ToArray();
                    }
                }
            }
            catch (Exception ex)
            {
                logger.WriteLog($"Error to Generate AVGStorageDays35 Report: {ex}");
                throw new CustomException((int)ErrorMessages.ErrorCodes.GlobalError, ex.Message);
            }
        }

        private byte[] GenerateSerialsReport(DataSet dataSet, ReportRequest request, ReportDtl reportDtl, Manifest manifest)
        {
            try
            {
                using (var workbook = new XLWorkbook())
                {
                    var worksheet = workbook.Worksheets.Add(reportDtl.ReportID);

                    PrintSettings(worksheet);

                    string columnName = "משתמש";

                    if (!dataSet.Tables[0].Columns.Contains(columnName))
                        dataSet.Tables[0].Columns.Add(columnName, request.User?.GetType() ?? typeof(string));

                    foreach (DataRow row in dataSet.Tables[0].Rows)
                    {
                        row[columnName] = request.User;
                    }

                    int currentRow = 1;
                    int numberOfColumns = dataSet.Tables[0].Columns.Count;

                    var table1 = worksheet.Cell(currentRow, 1).InsertTable(dataSet.Tables[0]);
                    ApplyTableStyleBoldHeadings(table1);


                    currentRow += dataSet.Tables[0].Rows.Count + 1;
                    worksheet.Cell(currentRow, 1).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Right;
                    worksheet.Range(currentRow, 1, currentRow, numberOfColumns).Style.Border.TopBorder = XLBorderStyleValues.Thin;

                    int startCalc = 2;

                    var columnsToCount = new List<string> { "מס' סריאלי" };

                    if (reportDtl.ReportID == ReportID.SerialsIn36.ToString())
                        columnsToCount.Add("תאריך כניסה");

                    else if (reportDtl.ReportID == ReportID.SerialsOut37.ToString())
                        columnsToCount.Add("תאריך מסירה");

                    ApplyColumnFormula(worksheet, dataSet.Tables[0], startCalc, currentRow, columnsToCount, "COUNTA");

                    ApplyNumberFormatToSheet(worksheet);

                    worksheet.Columns().AdjustToContents();

                    using (var stream = new MemoryStream())
                    {
                        workbook.SaveAs(stream);
                        return stream.ToArray();
                    }
                }
            }
            catch (Exception ex)
            {
                logger.WriteLog($"Error to Generate Serials Report: {ex}");
                throw new CustomException((int)ErrorMessages.ErrorCodes.GlobalError, ex.Message);
            }
        }

        private byte[] GenerateDelivery26Report(DataSet dataSet, ReportRequest request, ReportDtl reportDtl, Manifest manifest)
        {
            try
            {
                using (var workbook = new XLWorkbook())
                {
                    var worksheet = workbook.Worksheets.Add(reportDtl.ReportID);

                    PrintSettings(worksheet);

                    int currentRow = 1;
                    int numberOfColumns = dataSet.Tables[1].Columns.Count;

                    AddHeader2(worksheet, request, reportDtl, manifest, currentRow, numberOfColumns);

                    currentRow += 2;

                    StringBuilder filter = new StringBuilder();

                    if (request.Parameters.ContainsKey("BilledImporterID") && request.Parameters["BilledImporterID"] != null)
                    {
                        string billedImporterID = request.Parameters["BilledImporterID"]?.ToString();
                        if (!string.IsNullOrEmpty(billedImporterID))
                        {
                            string billedImporter = dataSet.Tables[2].Rows[0]["BilledImporter"].ToString();
                            if (filter.Length > 0) filter.Append(", ");
                            filter.Append($"ללקוח: {billedImporter} ({billedImporterID})");
                        }
                    }

                    if (request.Parameters.ContainsKey("FromGush") && request.Parameters["FromGush"] != null && request.Parameters.ContainsKey("ToGush") && request.Parameters["ToGush"] != null)
                    {
                        string fromGush = request.Parameters["FromGush"]?.ToString();
                        string toGush = request.Parameters["ToGush"]?.ToString();
                        if (!string.IsNullOrEmpty(fromGush) && !string.IsNullOrEmpty(toGush))
                        {
                            string FormattedFromGush = fromGush.Length > 2 ? $"{fromGush.Substring(0, 2)}/{fromGush.Substring(2)}" : fromGush;
                            string FormattedToGush = toGush.Length > 2 ? $"{toGush.Substring(0, 2)}/{toGush.Substring(2)}" : toGush;

                            if (filter.Length > 0) filter.Append(", ");
                            filter.Append($"מגוש {FormattedFromGush} עד גוש {FormattedToGush}");
                        }
                    }

                    worksheet.Cell(currentRow, 1).Value = filter.ToString();
                    worksheet.Range(worksheet.Cell(currentRow, 1), worksheet.Cell(currentRow, numberOfColumns)).Merge();
                    worksheet.Cell(currentRow, 1).Style.Font.Bold = true;
                    worksheet.Cell(currentRow, 1).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Right;

                    currentRow += 2;

                    DataTable info = dataSet.Tables[0]; // Gush table
                    DataTable details = dataSet.Tables[1]; // Detail table

                    var columnsToSum = new List<string> { "נפח במסירה", "משקל במסירה", "כמות במסירה" };

                    foreach (DataRow infoRow in info.Rows)
                    {
                        var gushInfo = info.Clone();
                        gushInfo.ImportRow(infoRow);

                        var table1 = worksheet.Cell(currentRow, 1).InsertTable(gushInfo);
                        ApplyTableStyleBoldHeadings(table1);
                        currentRow += gushInfo.Rows.Count + 2;

                        var gush = infoRow["גוש"];
                        var detailRows = details.AsEnumerable().Where(r => r["גוש"].Equals(gush));

                        if (detailRows.Any())
                        {
                            DataTable detailTable = details.Clone();

                            foreach (var dr in detailRows)
                                detailTable.ImportRow(dr);

                            detailTable.Columns.Remove("גוש");

                            int detailStartRow = currentRow;

                            var table2 = worksheet.Cell(currentRow, 1).InsertTable(detailTable);
                            ApplyTableStyleBoldHeadings(table2);

                            currentRow += detailTable.Rows.Count + 1;
                            worksheet.Cell(currentRow, 1).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Right;
                            worksheet.Range(currentRow, 1, currentRow, detailTable.Columns.Count).Style.Border.TopBorder = XLBorderStyleValues.Thin;

                            ApplyColumnFormula(worksheet, detailTable, detailStartRow, currentRow, columnsToSum, "SUM");
                            var mergeRange = worksheet.Range(currentRow, 1, currentRow, 3);
                            mergeRange.Merge();
                            mergeRange.Value = "סה\"כ מסירות בגוש:";
                            mergeRange.Style.Font.Bold = true;
                            currentRow++;

                        }

                        if (request.IsPrint)
                        {
                            worksheet.PageSetup.AddHorizontalPageBreak(currentRow - 1); // Page break
                        }
                        else
                        {
                            currentRow++;
                        }
                    }


                    ApplyNumberFormatToSheet(worksheet);

                    worksheet.Columns().AdjustToContents();


                    using (var stream = new MemoryStream())
                    {
                        workbook.SaveAs(stream);
                        return stream.ToArray();
                    }
                }
            }
            catch (Exception ex)
            {
                logger.WriteLog($"Error to Generate Delivery26 Report: {ex}");
                throw new CustomException((int)ErrorMessages.ErrorCodes.GlobalError, ex.Message);
            }
        }


    }
}
