using Google;
using GoogleSheetsToDocs.Models;
using GoogleSheetsToDocs.Services.Interfaces;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Logging;
using System.Diagnostics;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace GoogleSheetsToDocs.Controllers
{
    public class HomeController : Controller
    {
        private readonly ILogger<HomeController> _logger;
        private readonly IConvertService convertService;

        public HomeController(ILogger<HomeController> logger, IConvertService convertService)
        {
            _logger = logger;
            this.convertService = convertService;
        }

        public IActionResult Index()
        {
            SheetVM VM = new();

            return View(VM);
        }

        public async Task<IActionResult> Convert(string SheetId, string DocId)
        {
            string sheetId= Regex.Match(SheetId, @"[-\w]{25,}", RegexOptions.IgnoreCase).Value;
            string docId = Regex.Match(DocId, @"[-\w]{25,}", RegexOptions.IgnoreCase).Value;
            SheetVM VM = new()
            {
                SheetId =sheetId,
                DocId=docId
            };

            try
            {
                await convertService.ExtractDataFromSheet(VM);
                await convertService.CreateDoc(VM);
            }
            catch (GoogleApiException ex)
            {
                if (ex.HttpStatusCode == System.Net.HttpStatusCode.NotFound)
                    VM.Error = "Spreadsheet was not found";
                else
                    VM.Error = ex.Message;
            }

            return View("Index", VM);
        }

        [ResponseCache(Duration = 0, Location = ResponseCacheLocation.None, NoStore = true)]
        public IActionResult Error()
        {
            return View(new ErrorVM { RequestId = Activity.Current?.Id ?? HttpContext.TraceIdentifier });
        }
    }
}
