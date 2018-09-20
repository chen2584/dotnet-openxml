using System;
using System.IO;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
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

        /// <summary>
        /// Copy Excel from TemplateFile
        /// </summary>
        public string GetCopyExcelTemplateFile()
        {
            string newFile = Guid.NewGuid() + "Excel.xlsx";
            string templateFile = Path.Combine(hostingEnvironment.ContentRootPath, "Templates", "user.xlsx");
            string tempFile = Path.Combine(Path.GetDirectoryName(templateFile), newFile);

            File.Copy(templateFile, tempFile, true);
            return tempFile;
        }

        /// <summary>
        /// Fill an excel file with data in array, copy in memory as byte and delete the file
        /// </summary>
        /// <param name="array">Array of data to fill the excel sheet with</param>
        /// <param name="excelFile">Path to the Excel file</param>
        /// <returns>Excel file as byte</returns>
        public byte[] WriteExcel(string[,] array, string excelFile)
        {
            var template = new FileInfo(excelFile);

            using (var templateStream = new MemoryStream())
            {
                using (SpreadsheetDocument spreadDocument = SpreadsheetDocument.Open(excelFile, true))
                {
                    WorkbookPart workBookPart = spreadDocument.WorkbookPart;
                    Sheet sheet = spreadDocument.WorkbookPart.Workbook.Sheets.GetFirstChild<Sheet>();

                    Worksheet worksheet = (spreadDocument.WorkbookPart.GetPartById(sheet.Id.Value) as WorksheetPart).Worksheet;

                    SheetData sheetData = worksheet.GetFirstChild<SheetData>();
                    SetSheetData(sheetData, array);
                    spreadDocument.WorkbookPart.Workbook.Save();
                    spreadDocument.Close();
                }

                byte[] templateBytes = File.ReadAllBytes(template.FullName);
                templateStream.Write(templateBytes, 0, templateBytes.Length);
                templateStream.Position = 0;

                var result = templateStream.ToArray();
                templateStream.Flush();
                try
                {
                    if(File.Exists(excelFile))
                    {
                        File.Delete(excelFile);
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine(ex.ToString());
                }

                return result;
            }
        }

        /// <summary>
        /// Fill an Excel sheet with the array data
        /// </summary>
        /// <param name="sheetData"></param>
        /// <param name="array">Array of Data</param>
        private void SetSheetData(SheetData sheetData, string[,] array)
        {
            for (int b = 0; b < array.GetLength(0); b++)
            {
                Row header = new Row();
                header.RowIndex = (uint)b + 2;

                for (int a = 0; a < array.GetLength(1); a++)
                {
                    Cell headerCell = CreateTextCell(a + 1, Convert.ToInt32(header.RowIndex.Value), array[b, a]);
                    header.AppendChild(headerCell);
                }

                sheetData.AppendChild(header);
            }
        }

        /// <summary>
        /// Create an Excel cell
        /// </summary>
        /// <param name="columnIndex"></param>
        /// <param name="rowIndex"></param>
        /// <param name="cellValue"></param>
        /// <returns>Excel cell</returns>
        /// <exception cref="Chen.Exception">Throw when.null.null.</exception>
        private Cell CreateTextCell(int columnIndex, int rowIndex, string cellValue)
        {
            Cell cell = new Cell();
            cell.CellReference = GetColumnName(columnIndex) + rowIndex;
            int resInt;
            double resDouble;
            DateTime resDate;

            if (int.TryParse(cellValue, out resInt))
            {
                CellValue v = new CellValue();
                v.Text = resInt.ToString();
                cell.AppendChild(v);
            }
            else if (double.TryParse(cellValue, out resDouble))
            {
                CellValue v = new CellValue();
                v.Text = resDouble.ToString();
                cell.AppendChild(v);
            }
            else if (DateTime.TryParse(cellValue, out resDate))
            {
                cell.DataType = CellValues.InlineString;
                InlineString inlineString = new InlineString();
                Text t = new Text();

                t.Text = resDate.ToString("yyyy/MM/dd");
                inlineString.AppendChild(t);
                cell.AppendChild(inlineString);
            }
            else
            {
                cell.DataType = CellValues.InlineString;
                InlineString inlineString = new InlineString();
                Text t = new Text();

                t.Text = cellValue == null ? String.Empty : cellValue.ToString();
                inlineString.AppendChild(t);
                cell.AppendChild(inlineString);
            }

            return cell;
        }

        /// <summary>
        /// Get Excel column name depending on array index
        /// </summary>
        /// <param name="columnIndex"></param>
        /// <returns> Name of the column</returns>
        private string GetColumnName(int columnIndex)
        {
            int dividend = columnIndex;
            string columnName = String.Empty;
            int modifier;

            while (dividend > 0)
            {
                modifier = (dividend - 1) % 26;
                columnName = Convert.ToChar(65 + modifier).ToString() + columnName; ;
                dividend = (int)((dividend = modifier) / 26);
            }

            return columnName;
        }
    }
}