using DocumentFormat.OpenXml.Wordprocessing;
using Fingers10.ExcelExport.Attributes;
using System.ComponentModel.DataAnnotations;

namespace TestForFinger10
{
    public class WeatherForecast
    {
        public DateTime Date { get; set; }

        [IncludeInReport(Order = 1)]
        [Display(Name = "TemperatureC")]
        public int TemperatureC { get; set; }

        public int TemperatureF => 32 + (int)(TemperatureC / 0.5556);

        public string? Summary { get; set; }

        [NestedIncludeInReport]
        public City City { get; set; }
    }

    public class City
    {
        [IncludeInReport(Order = 2)]
        [Display(Name = "CityID")]
        public int Id { get; set; }
    }
}