using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using Google.Apis.Auth.OAuth2;
using Google.Apis.Services;
using Google.Apis.Sheets.v4;
using Google.Apis.Sheets.v4.Data;
using GoogleSheetsToDocs.Services.Interfaces;
using Microsoft.Extensions.Configuration;
using static Google.Apis.Sheets.v4.SpreadsheetsResource.ValuesResource.GetRequest;

namespace GoogleSheetsToDocs.Services.Implementations
{
    public class GoogleSheetsApiService : IGoogleSheetsApiService
    {
        private readonly SheetsService googleSheetsService;
        private readonly IGoogleAuthService googleAuthService;

        public GoogleSheetsApiService(IGoogleAuthService googleAuthService)
        {
            this.googleAuthService = googleAuthService;

            UserCredential credential = this.googleAuthService.GetUserCredential();

            googleSheetsService = new SheetsService(new BaseClientService.Initializer
            {
                ApplicationName = "GoogleSheetToDoc app",
                HttpClientInitializer = credential
            });
        }

        public async Task<IList<IList<object>>> GetSheetValuesRange(string SheetId, string range, ValueRenderOptionEnum? valueRenderOption = null)
        {
            var request = googleSheetsService.Spreadsheets.Values.Get(SheetId, range);
            request.ValueRenderOption = valueRenderOption;

            ValueRange response = await request.ExecuteAsync();
            return response.Values;
        }

        public async Task<Sheet> GetSheetRange(string SheetId, string range, bool includeGridData)
        {
            SpreadsheetsResource.GetRequest request = googleSheetsService.Spreadsheets.Get(SheetId);
            request.Ranges = new List<string> { range };
            request.IncludeGridData = includeGridData;

            Spreadsheet response = await request.ExecuteAsync();

            return response.Sheets.Count > 0 ? response.Sheets[0] : null;
        }
    }
}
