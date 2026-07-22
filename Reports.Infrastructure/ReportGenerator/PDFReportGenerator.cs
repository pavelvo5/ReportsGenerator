using DinkToPdf;
using PuppeteerSharp;
using PuppeteerSharp.Media;
using Reports.Infrastructure.DTOs;
using Reports.Infrastructure.Exceptions;
using Reports.Infrastructure.Logger;
using Reports.Infrastructure.Models;
using Reports.Infrastructure.Repositories;
using Stubble.Core.Builders;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing.Printing;
using System.IO;
using System.Linq;
using System.Printing;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using static Reports.Infrastructure.Models.Enums;

namespace Reports.Infrastructure.ReportGenerator
{
    public class PDFReportGenerator : IReportGenerator
    {
        public Enums.ReportType Type => Enums.ReportType.PDF;
        private readonly IReportRepositoryDapper repositoryDapper;
        private readonly ILogger logger;
        private readonly string sumatraPDF;
        private readonly string chrome;


        public PDFReportGenerator(IReportRepositoryDapper repositoryDapper, ILogger logger, string sumatraPDF, string chrome)
        {
            this.repositoryDapper = repositoryDapper;
            this.logger = logger;
            this.sumatraPDF = sumatraPDF;
            this.chrome = chrome;
        }

        public async Task<byte[]> ExecuteAsync(ReportRequest request, ReportDtl reportDtl)
        {
            try
            {
                string htmlContent = await BuildHtmlContentAsync(request, reportDtl);

                if (string.IsNullOrEmpty(htmlContent))
                {
                    logger.WriteLog($"No data was returned for the report.");
                    return null;
                }

                byte[] pdf = await HtmlToPdfDocument(htmlContent);

                if (request.IsPrint)
                {
                    PrintPdfDocument(pdf, request.PrinterName);
                }

                return pdf;

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


        private Dictionary<string, object> FlattenObject(object obj, string prefix = "")
        {
            try
            {
                if (obj == null) return null;

                var result = new Dictionary<string, object>();

                foreach (var prop in obj.GetType().GetProperties())
                {
                    var value = prop.GetValue(obj);

                    if (value == null)
                    {
                        result[prefix + prop.Name] = null;
                    }
                    else if (IsSimpleType(value.GetType()))
                    {
                        if (value is DateTime dt)
                        {
                            result[prefix + prop.Name] = dt.ToString("dd/MM/yyyy");
                        }
                        else
                        {
                            result[prefix + prop.Name] = value;
                        }
                    }
                    else if (value is IEnumerable list && value.GetType() != typeof(string))
                    {
                        // If the field is a list (including lists of objects)
                        var listData = new List<Dictionary<string, object>>();
                        foreach (var item in list)
                        {
                            listData.Add(FlattenObject(item, prefix + prop.Name + "_"));
                        }
                        result[prefix + prop.Name] = listData;
                    }
                    else
                    {
                        // If it's another internal object (not a list) - flatten the object inside
                        var flattenedObject = FlattenObject(value, prefix + prop.Name + "_");
                        foreach (var subProp in flattenedObject)
                        {
                            result[subProp.Key] = subProp.Value;
                        }
                    }
                }

                return result;
            }
            catch (Exception ex)
            {
                logger.WriteLog($"Error to Flatten Object to dictionary: {ex}");
                throw new CustomException((int)ErrorMessages.ErrorCodes.GlobalError, ex.Message);
            }
        }

        private bool IsSimpleType(Type type)
        {
            return type.IsPrimitive || type.IsValueType || type == typeof(string) || type == typeof(decimal) || type == typeof(DateTime);
        }

        private string GenerateHtml(Dictionary<string, object> data, ReportDtl reportDtl)
        {
            try
            {
                string basePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "bin", "Templates");

                string template = File.ReadAllText(Path.Combine(basePath, reportDtl.Template));

                var stubble = new StubbleBuilder().Build();


                if (!string.IsNullOrEmpty(reportDtl.HeaderTemplate))
                {
                    string headerTemplate = File.ReadAllText(Path.Combine(basePath, reportDtl.HeaderTemplate));
                    data["HeaderHtml"] = stubble.Render(headerTemplate, data);
                }

                if (!string.IsNullOrEmpty(reportDtl.TitleTemplate))
                {
                    string titleTemplate = File.ReadAllText(Path.Combine(basePath, reportDtl.TitleTemplate));
                    data["TitleHtml"] = stubble.Render(titleTemplate, data);
                }

                if (!string.IsNullOrEmpty(reportDtl.FooterTemplate))
                {
                    string footerTemplate = File.ReadAllText(Path.Combine(basePath, reportDtl.FooterTemplate));
                    data["FooterHtml"] = stubble.Render(footerTemplate, data);
                }


                string html = stubble.Render(template, data);

                logger.WriteLog("Generate HTML completed successfully.");

                return html;
            }
            catch (Exception ex)
            {
                logger.WriteLog($"Error to generate HTML: {ex}");
                throw new CustomException((int)ErrorMessages.ErrorCodes.GlobalError, ex.Message);
            }
        }
        private async Task<byte[]> HtmlToPdfDocument(string htmlContent)
        {
            try
            {
                var browser = await Puppeteer.LaunchAsync(new LaunchOptions
                {
                    Headless = true,
                    ExecutablePath = chrome
                });
                var page = await browser.NewPageAsync();

                await page.SetContentAsync(htmlContent);

                string margin = "12mm";

                var pdfOptions = new PdfOptions
                {
                    MarginOptions = new MarginOptions
                    {
                        Top = margin,
                        Bottom = margin,
                        Left = margin,
                        Right = margin
                    }
                };

                var pdfBuffer = await page.PdfDataAsync(pdfOptions);

                await browser.CloseAsync();

                logger.WriteLog("Generate PDF from HTML completed successfully.");

                return pdfBuffer;

            }
            catch (Exception ex)
            {
                logger.WriteLog($"Error to generate PDF from HTML: {ex}");
                throw new CustomException((int)ErrorMessages.ErrorCodes.GlobalError, ex.Message);
            }
        }
        private void PrintPdfDocument(byte[] pdf, string printerName)
        {
            string tempFilePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".pdf");

            try
            {
                // Saving the array to a temporary file
                File.WriteAllBytes(tempFilePath, pdf);

                PrintQueueStatus status = GetPrinterStatus(printerName);

                if (status.HasFlag(PrintQueueStatus.PaperProblem) ||
                status.HasFlag(PrintQueueStatus.Offline) ||
                status.HasFlag(PrintQueueStatus.Error))
                {
                    throw new CustomException((int)ErrorMessages.ErrorCodes.FailedToPrint, ErrorMessages.Messages[(int)ErrorMessages.ErrorCodes.FailedToPrint]);
                }

                // Creating parameters for running SumatraPDF for printing
                ProcessStartInfo psi = new ProcessStartInfo
                {
                    FileName = sumatraPDF,
                    Arguments = $"-print-to \"{printerName}\" -silent \"{tempFilePath}\"",
                    RedirectStandardError = true,
                    RedirectStandardOutput = true,
                    UseShellExecute = false,
                    CreateNoWindow = true
                };


                using (Process process = new Process { StartInfo = psi })
                {
                    process.Start();

                    string errorOutput = process.StandardError.ReadToEnd();
                    string standardOutput = process.StandardOutput.ReadToEnd();

                    process.WaitForExit();

                    if (process.ExitCode != 0)
                    {
                        logger.WriteLog($"Error to Print PDF Document. ExitCode: {process.ExitCode}");
                        if (!string.IsNullOrEmpty(errorOutput))
                        {
                            logger.WriteLog($"Error details: {errorOutput}");
                        }
                        throw new CustomException((int)ErrorMessages.ErrorCodes.FailedToPrint, ErrorMessages.Messages[(int)ErrorMessages.ErrorCodes.FailedToPrint]);
                    }
                    else
                    {
                        logger.WriteLog($"Print PDF Document completed successfully.");
                    }
                }
            }
            catch (CustomException ex)
            {
                logger.WriteLog($"Error to Print PDF Document: {ex}");
                throw;
            }
            catch (Exception ex)
            {
                logger.WriteLog($"Error to Print PDF Document: {ex}");
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

        private ReportRequest CloneRequestWithSplit(ReportRequest original, string splitKey, string splitValue)
        {
            var clone = new ReportRequest
            {
                ReportID = original.ReportID,
                ManifestID = original.ManifestID,
                Parameters = new Dictionary<string, object>(original.Parameters),
                IsPrint = original.IsPrint,
                PrinterName = original.PrinterName,
                User = original.User,
                ConnectionString = original.ConnectionString
            };

            clone.Parameters[splitKey] = splitValue;
            return clone;
        }

        private async Task<string> BuildInvRepForCustomsHtmlContentAsync(ReportRequest request, ReportDtl reportDtl, string splitValues)
        {
            try
            {
                var values = splitValues.Split(',').Select(v => v.Trim()).ToList();
                var htmlBuilder = new StringBuilder();

                for (int i = 0; i < values.Count; i++)
                {
                    var clonedRequest = CloneRequestWithSplit(request, "MultyGush", values[i]);

                    var dataReport = await repositoryDapper.GetDataAsync(clonedRequest, reportDtl);

                    Dictionary<string, object> data = FlattenObject(dataReport);

                    if (data != null && data.Any())
                    {
                        data["RequestingUser"] = clonedRequest.User;

                        string html = GenerateHtml(data, reportDtl);
                        htmlBuilder.Append(html);

                        if (i < values.Count - 1)
                            htmlBuilder.Append("<div class='page-break'></div>");
                    }
                }

                return htmlBuilder.ToString();
            }
            catch (CustomException ex)
            {
                logger.WriteLog($"Error to Build InvRepForCustoms Html Content Async: {ex}");
                throw;
            }
            catch (Exception ex)
            {
                logger.WriteLog($"Error to Build InvRepForCustoms Html Content Async: {ex}");
                throw new CustomException((int)ErrorMessages.ErrorCodes.GlobalError, ex.Message);
            }
        }

        private async Task<string> BuildHtmlContentAsync(ReportRequest request, ReportDtl reportDtl)
        {
            try
            {
                if (reportDtl.ReportID == ReportID.InvRepForCustoms.ToString() &&
                    request.Parameters.TryGetValue("MultyGush", out var splitValueObj) &&
                    splitValueObj is string splitValues &&
                    splitValues.Contains(","))
                {
                    return await BuildInvRepForCustomsHtmlContentAsync(request, reportDtl, splitValues);
                }
                else
                {
                    var dataReport = await repositoryDapper.GetDataAsync(request, reportDtl);
                    Dictionary<string, object> data = FlattenObject(dataReport);

                    if (data != null && data.Any())
                    {
                        // Not a database field - available to any PDF template as {{RequestingUser}}.
                        data["RequestingUser"] = request.User;

                        string htmlContent = GenerateHtml(data, reportDtl);
                        return htmlContent;
                    }
                    return null;
                }
            }
            catch (CustomException ex)
            {
                logger.WriteLog($"Error to Build Html Content Async: {ex}");
                throw;
            }
            catch (Exception ex)
            {
                logger.WriteLog($"Error to Build Html Content Async: {ex}");
                throw new CustomException((int)ErrorMessages.ErrorCodes.GlobalError, ex.Message);
            }
        }

    }
}
