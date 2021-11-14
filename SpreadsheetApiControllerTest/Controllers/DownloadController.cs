using ClosedXmlTest;
using Microsoft.AspNetCore.Mvc;
using SampleApiControllerTest;
using System.Text;

namespace SpreadsheetApiControllerTest.Controllers
{
    [ApiController]
    [Route("[controller]")]
    public class DownloadController : ControllerBase
    {
        private readonly ILogger<DownloadController> _logger;

        public DownloadController(ILogger<DownloadController> logger)
        {
            _logger = logger;
        }

        [HttpGet("OpenXmlDirectly")]
        public async Task<FileResult> GetUsingOpenXmlDirectly()
        {
            var excelStream = SampleExcelExporter.GetSampleExcelFile();
            excelStream.Position = 0;
            return File(excelStream, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "myspreadsheetusingopenxmldirectly.xlsx");
        }

        [HttpGet("ClosedXml")]
        public async Task<FileResult> GetUsingClosedXml()
        {
            var excelStream = SampleClosedXmlExcelExporter.GetSampleExcelFile();
            excelStream.Position = 0;
            return File(excelStream, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "myspreadsheetusingclosedxml.xlsx");
        }
    }
}