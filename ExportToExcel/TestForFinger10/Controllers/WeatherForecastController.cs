using Fingers10.ExcelExport.ActionResults;
using Microsoft.AspNetCore.Mvc;

namespace TestForFinger10.Controllers
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
                Summary = Summaries[Random.Shared.Next(Summaries.Length)],

                City = new City { Id = Random.Shared.Next(100,200)} 
            })
            .ToArray();
        }
        [HttpGet("ExcelFinger10")]
        public IActionResult Excel()
        {
            var data = Get();
            return new ExcelResult<WeatherForecast>(data, "sheet1", "ExcelFile.xlsx");
        }
    }
}