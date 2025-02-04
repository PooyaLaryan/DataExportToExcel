using ClosedXML.Excel;
using Microsoft.AspNetCore.Mvc;

namespace TestForCloseXml.Controllers
{
    [ApiController]
    [Route("[controller]")]
    public class WeatherForecastController : ControllerBase
    {
        private static readonly string[] Summaries = new[]
        {
        "Freezing", "Bracing", "Chilly", "Cool", "Mild", "Warm", "Balmy", "Hot", "Sweltering", "Scorching"
    };

        private readonly ILogger<WeatherForecastController> _logger;

        public WeatherForecastController(ILogger<WeatherForecastController> logger)
        {
            _logger = logger;
        }

        [HttpGet(Name = "GetWeatherForecast")]
        public IEnumerable<WeatherForecast> Get()
        {
            return Enumerable.Range(1, 5).Select(index => new WeatherForecast
            {
                Date = DateTime.Now.AddDays(index),
                TemperatureC = Random.Shared.Next(-20, 55),
                Summary = Summaries[Random.Shared.Next(Summaries.Length)]
            })
            .ToArray();
        }

        [HttpGet("Excel")]
        public void Excel()
        {
            var data = Get();
            using var workbook = new XLWorkbook();
            var workSheet = workbook.AddWorksheet();
            workSheet.Range(1, 1, 3, 3).SetValue(1);
            workbook.SaveAs(@"D:\test.xlsx");
        }

        [HttpGet("ExcelDataTable")]
        public void ExcelDataTable()
        {
            var data = Get().GetDataTable();
            using var workbook = new XLWorkbook();
            var workSheet = workbook.AddWorksheet();
            workSheet.Cell("A1").InsertData(data);  
            workbook.SaveAs("excel.xlsx");
        }

        [HttpGet("ExcelIEnumerable")]
        public void ExcelIEnumerable()
        {
            var data = Get();
            using var workbook = new XLWorkbook();
            var workSheet = workbook.AddWorksheet();
            workSheet.Cell("A1").InsertData(data);
            workbook.SaveAs(@"d:\excel.xlsx");
        }

        [HttpGet("ExcelStream")]
        public async Task<IActionResult> ExcelStream()
        {
            var data = Get();
            await using var stream = new MemoryStream();

            using var workbook = new XLWorkbook();
            var workSheet = workbook.AddWorksheet("Sheet1");
            workSheet.Cell("A1").InsertData(data);
            workbook.SaveAs(stream);
            stream.Position = 0;

            return File(stream.ToArray(), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "Test.xlsx");
        }
    }
}