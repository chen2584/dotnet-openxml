using System;
using System.IO;
using Microsoft.AspNetCore.Mvc;
using TestOpenXml.Services;

namespace TestOpenXml
{
    public class OpenXmlController : ControllerBase
    {
        private readonly ExcelService excelService;

        public OpenXmlController(ExcelService excelService)
        {
            this.excelService = excelService;
        }

        public FileResult Index()
        {
            string excelFile = excelService.GetCopyExcelTemplateFile();
            string[,] array = new string[,] {{"Chen", "MiddleChen", "SemapatChen", "20"}, {"Chen2", "MiddleChen2", "SemapatChen2", "20"}};
            var excelBytes = excelService.WriteExcel(array, excelFile);
            return File(excelBytes, "application/vnd.ms-excel", "excel.xlsx");
        }
    }
}