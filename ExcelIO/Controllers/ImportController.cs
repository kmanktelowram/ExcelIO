using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Web;
using System.Web.Http;
using System.Web.Http.Cors;
using System.Web.Mvc;
using ExcelIO.Models;
using GrapeCity.Spread.Sheets.ExcelIO;
using GrapeCity.Windows.SpreadSheet.Data;

namespace ExcelIO.Controllers
{
    [System.Web.Http.RoutePrefix("api/Import")]
    public class ImportController : ApiController
    {
        
        // POST api/Import
        [System.Web.Http.HttpPost]
        [System.Web.Http.Route("/")]
        public HttpResponseMessage Import()
        {
            HttpResponseMessage response;

            try
            {
                Importer importer = new Importer();
                string result = importer.ImportExcel(HttpContext.Current.Request.InputStream);
                importer = null;
                response = Request.CreateResponse(HttpStatusCode.OK, result);

                return response;
            }
            catch (Exception ex)
            {
                response = Request.CreateResponse(HttpStatusCode.InternalServerError, ex.Message);
                return response;
            }
        }
    }
}
