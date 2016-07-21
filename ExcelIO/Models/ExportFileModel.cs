using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using Newtonsoft.Json.Linq;

namespace ExcelIO.Models
{
    public class ExportFileModel
    {
        public JObject Spread { get; set; }
        public string ExportFileType { get; set; }
        public string ExportFileName { get; set; }
        public ExcelFileSetting Excel { get; set; }
        public PdfFileSetting PDF { get; set; }
    }

    public class ExcelFileSetting
    {
        public string SaveFlags { get; set; }
        public string Password { get; set; }
    }

    public class PdfFileSetting
    {
        public int[] SheetIndexes { get; set; }
        public JObject Setting { get; set; }
    }
}