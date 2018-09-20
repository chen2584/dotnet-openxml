using System;
using System.IO;
using Microsoft.AspNetCore.Hosting;

namespace TestOpenXml.Services
{
    public class ExcelService
    {
        readonly IHostingEnvironment hostingEnvironment;

        public ExcelService(IHostingEnvironment hostingEnvironment)
        {
            this.hostingEnvironment = hostingEnvironment;
        }

        public string GetCopyExcelTemplateFile()
        {
            string newFile = Guid.NewGuid() + "Excel.xlsx";
            string templateFile = Path.Combine(hostingEnvironment.ContentRootPath, "Templates", "user.xlsx");
            string tempFile = Path.Combine(Path.GetDirectoryName(templateFile), newFile);

            File.Copy(templateFile, tempFile, true); 
            return tempFile;
        }
    }
}