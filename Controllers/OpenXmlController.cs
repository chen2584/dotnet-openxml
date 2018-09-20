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
        public IActionResult Index()
        {
            return Ok(this.excelService.GetCopyExcelTemplateFile());
        }
    }
}