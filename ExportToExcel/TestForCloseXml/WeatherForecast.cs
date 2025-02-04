using static TestForCloseXml.MyAttribute;

namespace TestForCloseXml
{
    public class WeatherForecast
    {
        [ExcelColumn("date")]
        public DateTime Date { get; set; }

        [ExcelColumn("TemperatureC")]
        public int TemperatureC { get; set; }

        [ExcelColumn("TemperatureF")]
        public int TemperatureF => 32 + (int)(TemperatureC / 0.5556);

        [ExcelColumn("Summary")]
        public string? Summary { get; set; }
    }
}