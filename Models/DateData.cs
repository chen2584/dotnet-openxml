using System;
using System.ComponentModel.DataAnnotations;

namespace TestOpenXml
{
    public class DateData
    {
        [Range(typeof(DateTime), "1/2/2018", "1/2/2020", ErrorMessage = "StartDate is invalid")]
        public DateTime StartDate { get; set; }
        public DateTime EndDate { get; set; }
    }
}