using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Text.RegularExpressions;
using System.Web;
using System.Web.Http;
using System.Web.Http.Cors;
using System.Web.Mvc;
using ExcelIO.Models;
using GrapeCity.Spread.Sheets.ExcelIO;
using GrapeCity.Windows.SpreadSheet.Data;
using Microsoft.Ajax.Utilities;
using Newtonsoft.Json;

namespace ExcelIO.Controllers
{
    [System.Web.Http.RoutePrefix("api/Export")]
    public class ExportController: ApiController
    {
        // POST api/Export
        [System.Web.Http.HttpPost]
        [System.Web.Http.Route("/")]
        public HttpResponseMessage Export()
        {
            HttpResponseMessage response = new HttpResponseMessage();

            try
            {
                var postData = HttpContext.Current.Request.Form;
                var model = JsonConvert.DeserializeObject<ExportFileModel>(postData["data"]);

                if (model != null)
                {
                    var fileType = model.ExportFileType;
                    var exporter = new Exporter(model.Spread);
                    var stream = new MemoryStream();
                    string contentType = null;

                    switch (fileType)
                    {
                        case "xlsx":
                            contentType = ExportExcelFile(exporter, stream);
                            break;

                        case "xls":
                            contentType = ExportExcelFile(exporter, stream);
                            break;

                        default:
                            // not supported file types
                            response = Request.CreateResponse(HttpStatusCode.UnsupportedMediaType, "Invalid post data: file type not supported.");
                            return response;
                    }

                    exporter = null;

                    if (contentType != null)
                    {
                        stream.Seek(0, SeekOrigin.Begin);

                        string fileName = "";

                        string fileNameFromModel = GetFileName(model);

                        if (!fileNameFromModel.IsNullOrWhiteSpace())
                        {
                              fileName = Regex.Replace(fileNameFromModel, @"[^0-9a-zA-Z.]+", "");
                        }
                        
                        response.StatusCode = HttpStatusCode.OK;

                        response.Content = new StreamContent(stream);
                        response.Content.Headers.Add("Content-Disposition", "Attachment;filename=" + fileName);
                        response.Content.Headers.ContentType = new MediaTypeHeaderValue(contentType);

                        return response;
                    }
                }

                response = Request.CreateResponse(HttpStatusCode.InternalServerError, "Invalid post data.");
                return response;
            }
            catch (Exception ex)
            {
                response = Request.CreateResponse(HttpStatusCode.InternalServerError, ex.Message);
                return response;
            }
        }

        private static string ExportExcelFile(Exporter exporter, MemoryStream stream)
        {
            exporter.SaveExcel(stream);
            return "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
        }

        private static string GetFileName(ExportFileModel model)
        {
            if (model != null)
            {
                var fileType = model.ExportFileType;
                var extension = "." + fileType;
                var fileName = model.ExportFileName;

                if (string.IsNullOrEmpty(fileName))
                {
                    fileName = "export";
                }

                if (!fileName.EndsWith(extension))
                {
                    fileName += extension;
                }

                return fileName;
            }

            return string.Empty;
        }

    }
}